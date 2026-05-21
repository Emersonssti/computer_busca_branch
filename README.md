# Busca Branch Publicada — v4 (Computer Intelligence)

Aplicativo desktop para consultar branches remotas (`origin`) em que um arquivo foi alterado em um período, verificar integração e inspecionar o diff do commit. Versão **v4**.

## Usuário final (executável)

Distribua **`dist\BuscaBranchGit_v4.exe`**. Quem usa o programa **não precisa** de Python, pip nem acesso a este repositório.

### Requisitos na máquina do usuário

| Item | Obrigatório |
|------|-------------|
| Windows | Sim |
| **Git for Windows** ([download](https://git-scm.com/download/win)) | Sim — o `.exe` **não** embute o Git |
| Clone local do repositório analisado (ex.: ORACLE) | Sim |
| Rede / login | Só se marcar *Atualizar remoto (fetch)* no app |

### Como o usuário executa

1. Duplo clique em `BuscaBranchGit_v4.exe`.
2. Na abertura o programa **restaura a última pasta do clone** e **verifica login Git** (pergunta se deve abrir a janela Microsoft/Azure quando necessário).
3. **Pasta do clone** → raiz do repositório (Procurar…); fica salva para a próxima execução.
3. **Arquivo** → caminho do arquivo no projeto (ex.: `procedures/arquivo.sql`).
4. Deixar **Atualizar remoto (fetch)** desmarcado no dia a dia (evita login).
5. **Executar busca** → duplo clique na linha para ver alterações (diff).
6. **F1** — ajuda completa dentro do programa (linguagem para usuário leigo).

### Login Microsoft (Azure DevOps)

Se o `origin` for `dev.azure.com`, o fetch pode pedir conta **Microsoft corporativa** (não GitHub/Google). Login uma vez; depois o Windows costuma reutilizar a credencial.

---

## Desenvolvimento

### Stack

- Python 3.9+ · Tkinter · Git CLI
- Runtime: `openpyxl`, `Pillow` (`requirements.txt`)
- Build: PyInstaller 6+ (`requirements-build.txt`)

### Executar em modo desenvolvimento

```bat
git clone <url-do-repositorio>
cd computer_busca_branch
pip install -r requirements.txt
python index.py
```

### Gerar executável v4 (Windows)

```bat
cd computer_busca_branch
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
python -m PyInstaller BuscaBranchGit_v4.spec --noconfirm
```

Saída: **`dist\BuscaBranchGit_v4.exe`** (janela sem console). Inclui `logo.jpeg`, openpyxl e Pillow.

Build limpo (opcional):

```bat
rmdir /s /q build dist 2>nul
python -m PyInstaller BuscaBranchGit_v4.spec --noconfirm
```

### Estrutura principal

| Arquivo | Função |
|---------|--------|
| `index.py` | UI, busca Git, diff, Excel, ajuda |
| `BuscaBranchGit_v4.spec` | Empacotamento PyInstaller |
| `logo.jpeg` | Logo Computer Intelligence (obrigatório no build) |

Configuração do usuário: `%APPDATA%\BuscaBranchPublicada\settings.json` (`ultima_pasta_clone`, credencial já validada por repositório)

### Funcionalidades v4

- Interface com logo e tema Computer Intelligence
- Busca por período (60 dias padrão ou meses mm/aaaa)
- Colunas de integração (`integracao`, etc.)
- Diff do commit em janela interna (duplo clique / Ver alterações)
- Exportação Excel
- Fetch opcional (evita login repetido)
- Ajuda em abas (F1) voltada ao uso do `.exe`

### Notas técnicas

- PyInstaller não embute o executável `git`.
- Caminho padrão sugerido na UI: `Documents` do usuário.
- Repositórios Azure DevOps: autenticação via Git Credential Manager.

---

**Computer Intelligence** — Busca Branch Publicada v4
