import os
import re
import threading
import tkinter as tk
from calendar import monthrange
from datetime import date, datetime, timedelta, timezone
from tkinter import filedialog, messagebox, scrolledtext, ttk

from git import Repo
from git.exc import GitCommandError


def _caminho_git_no_repo(repo: Repo, arquivo_alvo: str) -> str:
    """Caminho relativo à raiz do repositório (formato que o Git espera)."""
    arquivo_alvo = arquivo_alvo.replace("\\", "/").strip()
    root = repo.working_tree_dir
    if not root:
        return arquivo_alvo
    root_norm = os.path.normpath(root)
    if os.path.isabs(arquivo_alvo):
        alvo_norm = os.path.normpath(arquivo_alvo)
        try:
            rel = os.path.relpath(alvo_norm, root_norm)
        except ValueError:
            return arquivo_alvo.replace("\\", "/")
        if rel.startswith(".."):
            return arquivo_alvo.replace("\\", "/")
        return rel.replace("\\", "/")
    return arquivo_alvo


def _intervalo_ultimos_dias(ndias: int) -> tuple[datetime, datetime]:
    """since = hoje - ndias 00:00; until exclusivo = amanhã 00:00 (inclui commits de hoje)."""
    today = date.today()
    since_dt = datetime.combine(today - timedelta(days=ndias), datetime.min.time())
    until_dt = datetime.combine(today + timedelta(days=1), datetime.min.time())
    return since_dt, until_dt


def _mm_yyyy_exibicao_padrao_60_dias() -> tuple[str, str]:
    """Texto inicial dos campos De/até: mês do (hoje − 60 dias) e mês de hoje."""
    today = date.today()
    ini = today - timedelta(days=60)
    return f"{ini.month}/{ini.year}", f"{today.month}/{today.year}"


def _normaliza_mm_yyyy_digitado(s: str) -> str | None:
    """Retorna 'm/aaaa' canônico ou None se inválido."""
    m = re.match(r"^\s*(\d{1,2})/(\d{4})\s*$", s.strip())
    if not m:
        return None
    mes, ano = int(m.group(1)), int(m.group(2))
    if not 1 <= mes <= 12:
        return None
    return f"{mes}/{ano}"


def _resolver_caminho_tracked_casefold(repo: Repo, caminho: str) -> str:
    """
    Resolve o caminho exatamente como está no índice Git, sem diferenciar maiúsculas/minúsculas.
    Aceita caminho relativo completo ou só o nome do arquivo (se for único no repo).
    """
    caminho = caminho.replace("\\", "/").strip()
    if not caminho:
        raise ValueError("Caminho do arquivo vazio.")

    tracked = [ln.strip() for ln in repo.git.ls_files().splitlines() if ln.strip()]
    if caminho in tracked:
        return caminho

    fold_map = {t.casefold(): t for t in tracked}
    if caminho.casefold() in fold_map:
        return fold_map[caminho.casefold()]

    base = caminho.rsplit("/", 1)[-1].casefold()
    hits = [t for t in tracked if t.rsplit("/", 1)[-1].casefold() == base]
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        amostra = "\n".join(hits[:15])
        mais = f"\n… e mais {len(hits) - 15}." if len(hits) > 15 else ""
        raise ValueError(
            f"Há {len(hits)} arquivos com esse nome (ignorando maiúsculas). "
            f"Informe o caminho completo no repositório, por exemplo:\n{amostra}{mais}"
        )

    raise ValueError(
        f"Arquivo não encontrado no repositório (nem por caminho nem por nome): {caminho!r}"
    )


def _parse_refs_integracao(texto: str) -> list[str]:
    """Lista refs pedidas (ex.: main → origin/main). Vírgula separa várias."""
    out: list[str] = []
    for part in texto.split(","):
        p = part.strip()
        if not p:
            continue
        if "/" not in p:
            p = f"origin/{p}"
        p = p.replace("\\", "/")
        out.append(p)
    return out


def _casar_ref_no_remoto(nomes_refs: set[str], pedido: str) -> str | None:
    """Encontra o nome exato da ref em origin (Git é sensível a maiúsculas no ref)."""
    pedido = pedido.replace("\\", "/").strip()
    if not pedido:
        return None
    if "/" not in pedido:
        candidato = f"origin/{pedido}"
    else:
        candidato = pedido
    alvo = candidato.casefold()
    for n in nomes_refs:
        if n.casefold() == alvo:
            return n
    return None


def _ref_padrao_origin(repo: Repo) -> str | None:
    """Branch padrão do clone (ex.: origin/integracao) via origin/HEAD."""
    try:
        out = repo.git.symbolic_ref("refs/remotes/origin/HEAD").strip()
    except GitCommandError:
        return None
    prefix = "refs/remotes/"
    if out.startswith(prefix):
        return out[len(prefix) :]
    return None


_FALLBACKS_INTEGRACAO = (
    "integracao",
    "INTEGRACAO",
    "homologacao",
    "main",
    "master",
    "develop",
    "release",
    "production",
)


def _resolver_refs_integracao_efetivas(
    repo: Repo, nomes_refs_origin: set[str], texto: str
) -> list[str]:
    """
    Refs que existem no remoto, respeitando maiúsculas reais do repo.
    Se o campo vier vazio ou nada casar, usa origin/HEAD e nomes comuns (integracao, main…).
    """
    resultado: list[str] = []
    vistos: set[str] = set()
    for pedido in _parse_refs_integracao(texto):
        casado = _casar_ref_no_remoto(nomes_refs_origin, pedido)
        if casado and casado not in vistos:
            vistos.add(casado)
            resultado.append(casado)
    if resultado:
        return resultado

    padrao = _ref_padrao_origin(repo)
    if padrao:
        casado = _casar_ref_no_remoto(nomes_refs_origin, padrao)
        if casado and casado.rsplit("/", 1)[-1] != "HEAD":
            return [casado]

    for fb in _FALLBACKS_INTEGRACAO:
        casado = _casar_ref_no_remoto(nomes_refs_origin, fb)
        if casado:
            return [casado]
    return []


_LABEL_SEM_REF_INTEGRACAO = "Sem branch de integração no remoto"


def _commit_ancestral_da_ref(repo: Repo, commit_sha: str, ref: str) -> bool:
    """True se commit_sha está no histórico que leva ao tip de ref (ex.: já integrado em produção)."""
    try:
        repo.git.merge_base("--is-ancestor", commit_sha, ref)
        return True
    except GitCommandError:
        return False


def _data_entrada_commit_na_ref(repo: Repo, commit_sha: str, ref: str) -> str | None:
    """
    Data (commit) em que a alteração passou a integrar o histórico da ref:
    primeiro commit no caminho ancestry-path entre commit_sha e o tip de ref (ex.: merge),
    ou o próprio commit_sha se não houver commits intermediários.
    """
    try:
        out = repo.git.rev_list(
            "--ancestry-path",
            f"{commit_sha}..{ref}",
            "--reverse",
            ref,
        ).strip()
    except GitCommandError:
        out = ""
    linhas = [ln.strip() for ln in out.splitlines() if ln.strip()]
    alvo_sha = linhas[0] if linhas else commit_sha
    try:
        return repo.commit(alvo_sha).committed_datetime.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return None


def parse_intervalo_mes_ano(inicio_mm_yyyy: str, fim_mm_yyyy: str) -> tuple[datetime, datetime]:
    """
    Interpreta mm/aaaa (intervalo inclusivo nos meses escolhidos).

    Retorna (since, until) para o Git: since = 1º dia do mês inicial 00:00;
    until = instante exclusivo = 1º dia do mês *seguinte* ao mês final 00:00
    (o Git usa --until como limite não inclusivo).
    """
    pat = re.compile(r"^\s*(\d{1,2})/(\d{4})\s*$")

    def um(s: str) -> tuple[int, int]:
        m = pat.match(s)
        if not m:
            raise ValueError(f"Formato inválido (use mm/aaaa): {s!r}")
        mes, ano = int(m.group(1)), int(m.group(2))
        if not 1 <= mes <= 12:
            raise ValueError(f"Mês inválido: {mes}")
        return mes, ano

    mi, yi = um(inicio_mm_yyyy)
    mf, yf = um(fim_mm_yyyy)
    inicio = datetime(yi, mi, 1, 0, 0, 0)
    if mf == 12:
        until_exclusivo = datetime(yf + 1, 1, 1, 0, 0, 0)
    else:
        until_exclusivo = datetime(yf, mf + 1, 1, 0, 0, 0)
    # Compara com o último momento do mês final (só para validar ordem)
    ultimo_dia = monthrange(yf, mf)[1]
    fim_inclusivo_visual = datetime(yf, mf, ultimo_dia, 23, 59, 59)
    if inicio > fim_inclusivo_visual:
        raise ValueError("Data inicial não pode ser depois da data final.")
    return inicio, until_exclusivo


def buscar_branches_com_alteracao(
    repo_path: str,
    arquivo_alvo: str,
    *,
    since: datetime | None = None,
    until: datetime | None = None,
    refs_integracao: str = "",
    progress_callback=None,
) -> list:
    """
    Branches em origin com pelo menos um commit que toca o arquivo.
    Se since/until forem informados, só commits nesse intervalo
    (since inclusivo; until = primeiro instante após o período, como o Git espera).

    refs_integracao: refs separadas por vírgula; o commit listado é considerado integrado
    se for ancestral do tip de alguma dessas refs (git merge-base --is-ancestor).
    """
    if not os.path.exists(repo_path):
        raise FileNotFoundError("Caminho do repositório inválido.")

    repo = Repo(repo_path)
    caminho = _caminho_git_no_repo(repo, arquivo_alvo)
    caminho = _resolver_caminho_tracked_casefold(repo, caminho)
    if progress_callback:
        progress_callback(f"Arquivo resolvido no repo: {caminho}")

    if "origin" not in [r.name for r in repo.remotes]:
        raise RuntimeError("Nenhum remote 'origin' configurado.")

    if progress_callback:
        progress_callback("Atualizando referências remotas (fetch)...")
    try:
        repo.remotes.origin.fetch()
    except GitCommandError as e:
        raise RuntimeError(f"Falha no fetch: {e}") from e

    nomes_refs_origin = {r.name for r in repo.remotes.origin.refs}
    refs_int_ok = _resolver_refs_integracao_efetivas(
        repo, nomes_refs_origin, refs_integracao or ""
    )
    if progress_callback:
        if not refs_int_ok:
            progress_callback(
                f"Aviso: {_LABEL_SEM_REF_INTEGRACAO}; colunas de integração não poderão comparar."
            )
        else:
            progress_callback(f"Comparando integração com: {', '.join(refs_int_ok)}")

    kwargs = {"paths": caminho, "max_count": 1}
    if since is not None:
        kwargs["since"] = since
    if until is not None:
        kwargs["until"] = until

    branches_com_modificacao = []
    refs = [r for r in repo.remotes.origin.refs if r.name.rsplit("/", 1)[-1] != "HEAD"]
    total = len(refs)

    for i, ref in enumerate(refs):
        if progress_callback:
            progress_callback(f"Analisando branches… {i + 1}/{total}")

        commits = list(repo.iter_commits(ref.name, **kwargs))

        if commits:
            c = commits[0]
            branches_com_modificacao.append(
                {
                    "branch": ref.name,
                    "ultimo_commit": c.hexsha[:7],
                    "data": c.committed_datetime.strftime("%Y-%m-%d %H:%M"),
                    "autor": c.author.name,
                    "_ts": c.committed_datetime,
                    "_sha": c.hexsha,
                }
            )

    def _chave_utc(dt: datetime) -> float:
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc).timestamp()
        return dt.astimezone(timezone.utc).timestamp()

    branches_com_modificacao.sort(key=lambda r: _chave_utc(r["_ts"]), reverse=True)
    for item in branches_com_modificacao:
        del item["_ts"]

    for item in branches_com_modificacao:
        sha = item.pop("_sha")
        if refs_int_ok:
            ref_match = None
            for rref in refs_int_ok:
                if _commit_ancestral_da_ref(repo, sha, rref):
                    ref_match = rref
                    break
            if ref_match:
                curto = ref_match.split("/", 1)[-1] if "/" in ref_match else ref_match
                item["integracao"] = f"Sim ({curto})"
                item["data_integracao"] = (
                    _data_entrada_commit_na_ref(repo, sha, ref_match) or "—"
                )
            else:
                item["integracao"] = "Não"
                item["data_integracao"] = "—"
        else:
            item["integracao"] = _LABEL_SEM_REF_INTEGRACAO
            item["data_integracao"] = "—"

    return branches_com_modificacao


_APP_TITLE = "Computer - Busca Branch Publicada"

_TEXTO_AJUDA = """
O QUE ESTA FERRAMENTA FAZ
• Consulta o remoto origin no seu clone local e lista branches publicadas em que o arquivo
  escolhido foi alterado no período indicado.
• Ordena do commit mais recente para o mais antigo.
• Indica se o commit encontrado já aparece no histórico da branch de integração (produção).

PASTA DO CLONE
• Caminho da raiz de um repositório Git já clonado na sua máquina.
• O aplicativo executa fetch no origin; é preciso Git instalado e permissão de acesso ao servidor.

ARQUIVO
• Caminho do arquivo em relação à raiz do repositório (ex.: procedures/arquivo.sql) ou caminho
  absoluto dentro do clone.
• Maiúsculas e minúsculas no nome/caminho são ignoradas; o programa localiza o arquivo real no Git.
• Se informar só o nome do arquivo e existir mais de um com o mesmo nome em pastas diferentes,
  será pedido o caminho completo.

PERÍODO (MM/AAAA)
• Os campos já abrem preenchidos com o mês aproximado de “há 60 dias” e o mês atual, para
  deixar claro qual é o período padrão.
• Enquanto você não mudar esse par de valores, a busca usa exatamente os últimos 60 dias
  (do calendário, não só meses inteiros).
• Se você alterar “De” ou “até”, a busca passa a usar meses completos: do primeiro dia do
  mês inicial ao último dia do mês final (ambos inclusivos).
• Se apagar os dois campos, a busca volta a usar só os últimos 60 dias (equivalente ao padrão).

BRANCH DE INTEGRAÇÃO
• Branch (ou branches, separadas por vírgula) que representa integração/produção no seu fluxo.
• Valor sugerido por padrão: integracao (vira origin/integracao no remoto).
• Se deixar em branco, o programa tenta origin/HEAD e nomes comuns (integracao, main, etc.).
• Colunas “Na integração” e “Data na integração” na grade:
  – Sim (nome): o commit listado já está no histórico dessa branch → costuma indicar que a
    alteração já foi integrada nessa linha.
  – Data na integração: data do commit em que essa alteração entrou no histórico da branch de
    integração (ex.: merge mais próximo no caminho até o tip; se o commit já está direto na
    linha, mostra a data desse commit).
  – Não: o commit ainda não aparece no histórico da branch de integração informada (sem data).
  – “Sem branch de integração no remoto”: nenhuma ref válida foi encontrada no origin para comparar.

RESULTADOS
• Copiar o nome da branch: botão acima da tabela, duplo clique na linha, Ctrl+C com foco na
  tabela, ou clique direito → Copiar.

ATALHOS
• F1 — abre esta janela de ajuda.

REQUISITOS NA MÁQUINA
• Git no PATH.
• Clone atualizado ou com rede para o fetch; credenciais são as já configuradas no seu ambiente
  (o executável não armazena usuário ou senha).
""".strip()


def _criar_app():
    default_repo = r"C:\Users\Catalog Locação\Documents\ORACLE"
    _de_padrao, _ate_padrao = _mm_yyyy_exibicao_padrao_60_dias()
    _de_padrao_key = _normaliza_mm_yyyy_digitado(_de_padrao)
    _ate_padrao_key = _normaliza_mm_yyyy_digitado(_ate_padrao)

    root = tk.Tk()
    root.title(_APP_TITLE)
    root.minsize(760, 520)
    root.geometry("1100x580")

    frm = ttk.Frame(root, padding=(16, 14))
    frm.grid(row=0, column=0, sticky="nsew")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    def mostrar_ajuda():
        win = tk.Toplevel(root)
        win.title(f"Ajuda — {_APP_TITLE}")
        win.minsize(520, 420)
        win.geometry("580x500")
        win.transient(root)
        hdr_a = ttk.Frame(win, padding=(12, 12, 12, 4))
        hdr_a.pack(fill=tk.X)
        ttk.Label(
            hdr_a,
            text=_APP_TITLE,
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            hdr_a,
            text="Leia abaixo ou role a página para ver todos os tópicos.",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))
        body = scrolledtext.ScrolledText(
            win,
            wrap=tk.WORD,
            width=72,
            height=22,
            font=("Segoe UI", 10),
            padx=10,
            pady=8,
        )
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))
        body.insert("1.0", _TEXTO_AJUDA)
        body.configure(state="disabled")
        ttk.Button(win, text="Fechar", command=win.destroy).pack(pady=(0, 14))

    hdr = ttk.Frame(frm)
    hdr.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
    hdr.columnconfigure(0, weight=1)
    ttk.Label(
        hdr,
        text=_APP_TITLE,
        font=("Segoe UI", 13, "bold"),
    ).grid(row=0, column=0, sticky="w")
    ttk.Button(hdr, text="Ajuda…", command=mostrar_ajuda, width=10).grid(
        row=0, column=1, sticky="e", padx=(12, 0)
    )

    crit = ttk.LabelFrame(frm, text="Critérios de busca", padding=(12, 10))
    crit.grid(row=1, column=0, columnspan=3, sticky="ew")
    crit.columnconfigure(1, weight=1)

    ttk.Label(crit, text="Pasta do clone").grid(row=0, column=0, sticky="w", pady=4)
    repo_var = tk.StringVar(value=default_repo)
    ent_repo = ttk.Entry(crit, textvariable=repo_var)
    ent_repo.grid(row=0, column=1, sticky="ew", padx=(10, 6), pady=4)

    def escolher_pasta():
        p = filedialog.askdirectory(title="Selecionar pasta do repositório Git", parent=root)
        if p:
            repo_var.set(p)

    ttk.Button(crit, text="Procurar…", command=escolher_pasta).grid(row=0, column=2, pady=4)

    ttk.Label(crit, text="Arquivo").grid(row=1, column=0, sticky="nw", pady=4)
    arquivo_var = tk.StringVar(value="procedures/TS.ASS_VALIDA_PF_ddl.sql")
    ttk.Entry(crit, textvariable=arquivo_var).grid(
        row=1, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=4
    )

    ttk.Label(crit, text="Período").grid(row=2, column=0, sticky="nw", pady=4)
    periodo_fr = ttk.Frame(crit)
    periodo_fr.grid(row=2, column=1, columnspan=2, sticky="w", padx=(10, 0), pady=4)
    inicio_var = tk.StringVar(value=_de_padrao)
    fim_var = tk.StringVar(value=_ate_padrao)
    ttk.Label(periodo_fr, text="De").pack(side=tk.LEFT)
    ttk.Entry(periodo_fr, textvariable=inicio_var, width=12).pack(side=tk.LEFT, padx=(6, 16))
    ttk.Label(periodo_fr, text="até").pack(side=tk.LEFT)
    ttk.Entry(periodo_fr, textvariable=fim_var, width=12).pack(side=tk.LEFT, padx=(6, 0))

    ttk.Label(crit, text="Branch de integração").grid(row=3, column=0, sticky="nw", pady=4)
    integracao_var = tk.StringVar(value="integracao")
    ttk.Entry(crit, textvariable=integracao_var).grid(
        row=3, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=4
    )

    acoes = ttk.Frame(frm)
    acoes.grid(row=2, column=0, columnspan=3, sticky="w", pady=(14, 8))
    btn_run = ttk.Button(acoes, text="Executar busca")
    btn_run.pack(side=tk.LEFT)

    status_var = tk.StringVar(value="Pronto.")
    ttk.Label(frm, textvariable=status_var).grid(row=3, column=0, columnspan=3, sticky="w")

    resultados = ttk.LabelFrame(frm, text="Resultados", padding=(10, 8))
    resultados.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
    frm.rowconfigure(4, weight=1)
    resultados.columnconfigure(0, weight=1)
    resultados.rowconfigure(1, weight=1)

    bar_resultados = ttk.Frame(resultados)
    bar_resultados.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    bar_resultados.columnconfigure(1, weight=1)

    tree_frame = ttk.Frame(resultados)
    tree_frame.grid(row=1, column=0, sticky="nsew")
    tree_frame.rowconfigure(1, weight=1)
    tree_frame.columnconfigure(0, weight=1)

    cols = ("branch", "commit", "data", "autor", "integracao", "data_integracao")
    tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
    tree.heading("branch", text="Branch publicada")
    tree.heading("commit", text="Commit")
    tree.heading("data", text="Data do commit")
    tree.heading("autor", text="Autor")
    tree.heading("integracao", text="Na integração")
    tree.heading("data_integracao", text="Data na integração")
    tree.column("branch", width=260, minwidth=120)
    tree.column("commit", width=72, minwidth=60)
    tree.column("data", width=120, minwidth=100)
    tree.column("autor", width=140, minwidth=80)
    tree.column("integracao", width=200, minwidth=90)
    tree.column("data_integracao", width=130, minwidth=100)

    sy = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    sx = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
    tree.grid(row=1, column=0, sticky="nsew")
    sy.grid(row=1, column=1, sticky="ns")
    sx.grid(row=2, column=0, sticky="ew")
    menu_ctx = tk.Menu(root, tearoff=0)

    root.bind("<F1>", lambda _e: mostrar_ajuda())

    def branch_da_selecao():
        sel = tree.selection()
        if not sel:
            return None
        vals = tree.item(sel[0], "values")
        return vals[0] if vals else None

    def copiar_branch_selecionada():
        b = branch_da_selecao()
        if not b:
            status_var.set("Selecione uma linha na tabela.")
            return
        root.clipboard_clear()
        root.clipboard_append(b)
        root.update_idletasks()
        curto = b if len(b) <= 72 else b[:69] + "…"
        status_var.set(f"Copiado: {curto}")

    menu_ctx.add_command(label="Copiar nome da branch", command=copiar_branch_selecionada)

    def menu_contexto(event):
        iid = tree.identify_row(event.y)
        if not iid:
            return
        tree.selection_set(iid)
        try:
            menu_ctx.tk_popup(event.x_root, event.y_root)
        finally:
            menu_ctx.grab_release()

    def on_ctrl_c(_event):
        copiar_branch_selecionada()
        return "break"

    tree.bind("<Button-3>", menu_contexto)
    tree.bind("<Control-c>", on_ctrl_c)
    tree.bind("<Control-C>", on_ctrl_c)

    def duplo_clique(_event):
        iid = tree.identify_row(_event.y)
        if iid:
            tree.selection_set(iid)
            copiar_branch_selecionada()

    tree.bind("<Double-1>", duplo_clique)

    ttk.Button(
        bar_resultados,
        text="Copiar nome da branch",
        command=copiar_branch_selecionada,
    ).grid(row=0, column=0, sticky="w")

    def set_busy(busy: bool):
        btn_run.state(["disabled"] if busy else ["!disabled"])

    def on_progress(msg: str):
        root.after(0, lambda m=msg: status_var.set(m))

    def executar():
        tree.delete(*tree.get_children())
        path = repo_var.get().strip()
        arq = arquivo_var.get().strip()
        di, df = inicio_var.get().strip(), fim_var.get().strip()

        if not path or not arq:
            messagebox.showwarning(
                _APP_TITLE,
                "Informe a pasta do clone e o arquivo.",
                parent=root,
            )
            return

        if not di and not df:
            since_dt, until_dt = _intervalo_ultimos_dias(60)
        elif not di or not df:
            messagebox.showwarning(
                _APP_TITLE,
                "Preencha “De” e “até” com mm/aaaa, ou apague os dois para usar os últimos 60 dias.",
                parent=root,
            )
            return
        else:
            dk = _normaliza_mm_yyyy_digitado(di)
            fk = _normaliza_mm_yyyy_digitado(df)
            if dk is None or fk is None:
                messagebox.showerror(
                    _APP_TITLE,
                    "Formato inválido. Use mm/aaaa em “De” e “até” (ex.: 3/2026).",
                    parent=root,
                )
                return
            if dk == _de_padrao_key and fk == _ate_padrao_key:
                since_dt, until_dt = _intervalo_ultimos_dias(60)
            else:
                try:
                    since_dt, until_dt = parse_intervalo_mes_ano(di, df)
                except ValueError as e:
                    messagebox.showerror(_APP_TITLE, str(e), parent=root)
                    return

        set_busy(True)
        status_var.set("Buscando…")

        def worker():
            err = None
            res = []
            try:
                res = buscar_branches_com_alteracao(
                    path,
                    arq,
                    since=since_dt,
                    until=until_dt,
                    refs_integracao=integracao_var.get().strip(),
                    progress_callback=on_progress,
                )
            except Exception as e:
                err = e

            def finish():
                set_busy(False)
                if err:
                    messagebox.showerror(_APP_TITLE, str(err), parent=root)
                    status_var.set("Erro.")
                    return
                for r in res:
                    tree.insert(
                        "",
                        "end",
                        values=(
                            r["branch"],
                            r["ultimo_commit"],
                            r["data"],
                            r["autor"],
                            r["integracao"],
                            r["data_integracao"],
                        ),
                    )
                status_var.set(f"Concluído: {len(res)} branch(es).")

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    btn_run.configure(command=executar)

    return root


if __name__ == "__main__":
    app = _criar_app()
    app.mainloop()
