/* TARGET SSW 475 — Central de Controle (fala com a API do server.py) */
const $ = (id) => document.getElementById(id);
const state = { regs: [], concluidos: {}, cursor: 0, running: false, lastCompra: null };

function tlog(msg, cls = "") {
  const t = $("terminal");
  const d = document.createElement("div");
  if (cls) d.className = cls;
  d.textContent = msg;
  t.appendChild(d);
  while (t.children.length > 800) t.removeChild(t.firstChild);
  t.scrollTop = t.scrollHeight;
}

async function api(path, opts) {
  const r = await fetch(path, opts);
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.erro || ("HTTP " + r.status));
  return j;
}

/* ---------- status + logs (polling) ---------- */
async function refreshStatus() {
  try {
    const s = await api("/api/status");
    state.running = s.running;
    const dot = $("statusDot"), txt = $("statusText");
    dot.className = "dot" + (s.running ? " run" : (s.exit_code && s.exit_code !== 0 ? " err" : " on"));
    txt.textContent = s.running ? `Rodando (PID ${s.pid})` : (s.exit_code && s.exit_code !== 0 ? `Parado (exit ${s.exit_code})` : "Pronto");
    $("btnStart").disabled = s.running;
    $("btnStart2").disabled = s.running;
    $("btnStop").disabled = !s.running;
    $("btnStop2").disabled = !s.running;
    $("execInfo").textContent = s.running
      ? `PID ${s.pid} · início ${s.started_at || "?"} · ${s.planilha_nome} · modo ${s.modo || "?"}`
      : (s.exit_code !== null && s.exit_code !== undefined ? `Última execução: exit ${s.exit_code}` : `Parado · planilha: ${s.planilha_nome || "—"}`);
    $("footStatus").textContent = txt.textContent;
    if (s.planilha_nome) {
      $("planilhaNome").textContent = s.planilha_nome;
      const c = s.credenciais;
      if (c.dominio && !$("inDominio").value) $("inDominio").value = c.dominio;
      if (c.usuario && !$("inUsuario").value) $("inUsuario").value = c.usuario;
      $("loginHint").textContent = c.configurado
        ? `Acesso salvo: ${c.dominio}/${c.usuario} (CPF ${c.cpf_masc})${c.senha_salva ? " · senha salva" : (c.senha_memoria ? " · senha na memória" : " · senha será pedida")}.`
        : "Nenhum acesso salvo — preencha e clique em Salvar acesso.";
    }
  } catch (e) { $("statusText").textContent = "Servidor fora?"; }
}

async function pollLogs() {
  try {
    const j = await api("/api/logs?cursor=" + state.cursor);
    state.cursor = j.cursor;
    for (const l of j.lines) {
      const t = l.txt;
      let cls = "";
      if (/\[OK\]|\[SUCESSO\]|gravada/i.test(t)) cls = "ok";
      else if (/\[ERRO\]|recusou|falha/i.test(t)) cls = "err";
      else if (/\[AVISO|LOGIN|WARN/i.test(t)) cls = "warn";
      else if (/^\[WEB\]/.test(t)) cls = "web";
      tlog(t, cls);
      // rastreia progresso por compra p/ pintar a tabela
      let m = t.match(/Compra (\w+)/);
      if (m && /Processando/.test(t)) { state.lastCompra = m[1]; paintRows(); }
      m = t.match(/Linha \d+ gravada! Lancamento: (\S+)/);
      if (m && state.lastCompra) {
        const n = state.lastCompra;
        state.concluidos[n] = m[1];
        state.lastCompra = null;
        paintRows(); updateStats();
      }
    }
  } catch (e) { /* servidor reiniciando */ }
}

/* ---------- planilha ---------- */
async function loadPlanilha() {
  $("tbody").innerHTML = `<tr><td colspan="17"><div class="empty">Lendo planilha…</div></td></tr>`;
  try {
    const j = await api("/api/planilha");
    state.regs = j.registros || [];
    state.concluidos = j.concluidos || {};
    $("planilhaNome").textContent = j.arquivo || "?";
    $("planilhaInfo").textContent = `${j.total} registros · ${Object.keys(state.concluidos).length} já lançados`;
    tlog(`[WEB] Planilha ${j.arquivo}: ${j.total} registros.`, "web");
  } catch (e) {
    $("tbody").innerHTML = `<tr><td colspan="17"><div class="empty">Erro: ${e.message}</div></td></tr>`;
    tlog("[WEB] Erro ao ler planilha: " + e.message, "err");
    return;
  }
  paintRows(); updateStats();
}

function rowStatus(r) {
  if (state.concluidos[r.ncompra]) return "success";
  if (state.lastCompra === r.ncompra && state.running) return "running";
  return "pending";
}

function paintRows() {
  const tb = $("tbody");
  if (!state.regs.length) { tb.innerHTML = `<tr><td colspan="17"><div class="empty">Sem dados — envie o .xlsx.</div></td></tr>`; return; }
  tb.innerHTML = state.regs.map((r, i) => {
    const st = rowStatus(r);
    const badge = st === "success" ? `<span class="pill p-success">Lançado</span>`
      : st === "running" ? `<span class="pill p-running">Rodando</span>` : `<span class="pill p-pending">Pendente</span>`;
    return `<tr class="${st === "success" ? "r-success" : st === "running" ? "r-running" : ""}">
      <td>${i + 1}</td><td>${badge}</td><td class="lanc">${state.concluidos[r.ncompra] || "-"}</td>
      <td>${r.unidade}</td><td>${r.fornecedor}</td><td>${r.evento}</td><td>${r.serie}</td>
      <td><strong>${r.ncompra}</strong></td><td>${r.modelo}</td><td>${r.valor}</td>
      <td>${r.emissao}</td><td>${r.entrada}</td><td>${r.vencimento}</td><td>${r.pagamento}</td>
      <td>${r.competencia}</td><td>${r.vlr_parcela}</td><td>${r.historico}</td></tr>`;
  }).join("");
  $("rowBadge").textContent = `${state.regs.length} linhas`;
}

function updateStats() {
  const total = state.regs.length;
  const done = state.regs.filter((r) => state.concluidos[r.ncompra]).length;
  $("statTotal").textContent = total;
  $("statSuccess").textContent = done;
  $("statPending").textContent = total - done;
}

/* ---------- ações ---------- */
async function salvarLogin() {
  const body = {
    dominio: $("inDominio").value, cpf: $("inCpf").value,
    usuario: $("inUsuario").value, senha: $("inSenha").value,
    salvar_senha: $("inSalvarSenha").checked,
  };
  try {
    await api("/api/credentials", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    $("inSenha").value = "";
    tlog("[WEB] Acesso salvo. Pode iniciar a automação.", "web");
    refreshStatus();
  } catch (e) { tlog("[WEB] Erro ao salvar acesso: " + e.message, "err"); alert(e.message); }
}

async function upload(e) {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  const fd = new FormData();
  fd.append("arquivo", f);
  try {
    const j = await api("/api/upload", { method: "POST", body: fd });
    tlog("[WEB] Planilha ativa: " + j.arquivo, "web");
    loadPlanilha(); refreshStatus();
  } catch (err) { tlog("[WEB] Upload falhou: " + err.message, "err"); alert(err.message); }
}

async function start() {
  const limite = parseInt($("inLimite").value || "0", 10) || 0;
  const visivel = $("inVisivel").checked;
  if (!state.regs.length) await loadPlanilha();
  try {
    const j = await api("/api/start", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ limite, visivel }) });
    tlog(`[WEB] Automação iniciada (PID ${j.pid}, modo ${visivel ? "visível" : "oculto"}).` + (visivel ? " O navegador vai abrir." : " Sem janela — acompanhe pelo terminal."), "web");
    refreshStatus();
  } catch (e) { tlog("[WEB] " + e.message, "err"); alert(e.message); }
}

async function stop() {
  try { await api("/api/stop", { method: "POST" }); tlog("[WEB] Parada solicitada.", "warn"); }
  catch (e) { tlog("[WEB] " + e.message, "err"); }
  refreshStatus();
}

/* ---------- colagem do Excel ---------- */
function contarColagem() {
  const v = $("pasteArea").value || "";
  const n = v.split(/\r\n|\n|\r/).filter((l) => l.trim()).length;
  $("pasteBadge").textContent = `${n} linha${n === 1 ? "" : "s"}`;
}

async function usarColagem() {
  const texto = $("pasteArea").value || "";
  if (!texto.trim()) return alert("Cole as linhas do Excel na caixa primeiro.");
  try {
    const j = await api("/api/colagem", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ texto }) });
    state.regs = j.registros || [];
    state.concluidos = j.concluidos || {};
    state.lastCompra = null;
    $("planilhaNome").textContent = j.arquivo;
    $("planilhaInfo").textContent = `${j.total} registros colados · ${Object.keys(state.concluidos).length} já lançados`;
    paintRows(); updateStats(); refreshStatus();
    tlog(`[WEB] Base atualizada da colagem: ${j.total} linhas.`, "web");
  } catch (e) { tlog("[WEB] Colagem rejeitada: " + e.message, "err"); alert(e.message); }
}

function exportCSV() {  if (!state.regs.length) return alert("Sem dados.");
  const H = ["Linha", "Status", "Lancamento", "Unidade", "Fornecedor", "Evento", "Serie", "N_Compra", "Modelo", "Valor", "Emissao", "Entrada", "Vencimento", "Pagamento", "Competencia", "Vlr_Parcela", "Historico"];
  const q = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const rows = state.regs.map((r, i) => [i + 1, rowStatus(r), state.concluidos[r.ncompra] || "", r.unidade, r.fornecedor, r.evento, r.serie, r.ncompra, r.modelo, r.valor, r.emissao, r.entrada, r.vencimento, r.pagamento, r.competencia, r.vlr_parcela, q(r.historico)].join(";"));
  const blob = new Blob(["﻿" + H.join(";") + "\r\n" + rows.join("\r\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "relatorio_ssw475.csv";
  a.click();
  tlog("[WEB] CSV exportado.", "web");
}

/* ---------- init ---------- */
document.addEventListener("DOMContentLoaded", () => {
  $("btnSalvarLogin").onclick = salvarLogin;
  $("btnUpload").onclick = () => $("fileInput").click();
  $("fileInput").onchange = upload;
  $("btnRecarregar").onclick = loadPlanilha;
  $("btnStart").onclick = start;
  $("btnStart2").onclick = start;
  $("btnStop").onclick = stop;
  $("btnStop2").onclick = stop;
  $("btnExport").onclick = exportCSV;
  $("btnClearLog").onclick = () => { $("terminal").innerHTML = ""; };
  $("pasteArea").addEventListener("input", contarColagem);
  $("btnUsarColagem").onclick = usarColagem;
  $("btnLimparColagem").onclick = () => { $("pasteArea").value = ""; contarColagem(); };
  tlog("Central TARGET SSW 475 pronta. Salvando o acesso e clicando em Iniciar, o robô faz tudo sozinho.", "web");
  refreshStatus(); loadPlanilha();
  setInterval(refreshStatus, 3000);
  setInterval(pollLogs, 1500);
});
