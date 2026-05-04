# Busca Branch Git (Computer)

Aplicativo em **Python + Tkinter** que consulta o remoto `origin`, faz `fetch` e lista branches em que um arquivo foi alterado no período, com indicação de integração.

## Requisitos em qualquer máquina

- **Python 3.9+** com **Tkinter** (no instalador oficial do Windows, mantenha *tcl/tk and IDLE* marcado).
- **Git** no PATH ([Git for Windows](https://git-scm.com/download/win) no Windows).
- **Não** é necessário `pip install` para rodar o programa (só usa a biblioteca padrão + o executável `git`).

## Clonar e executar (Windows)

```bat
git clone <url-do-repositorio>
cd computer_busca_branch
python index.py
```

Ou use o interpretador que preferir, desde que seja o mesmo que tem Tkinter.

## Clonar e executar (macOS)

```bash
git clone <url-do-repositorio>
cd computer_busca_branch
./run_macos.sh
```

Se o Tk acusar incompatibilidade de versão do sistema, instale Python/Tk pelo Homebrew conforme a mensagem do script.

## Gerar executável (.exe) no Windows

1. Instale Python 3.9+ e Git (como acima).
2. Na pasta do projeto:

   ```bat
   build_windows.bat
   ```

   Ou manualmente:

   ```bat
   python -m pip install -r requirements-build.txt
   python -m PyInstaller BuscaBranchGit_v3.spec
   ```

3. O arquivo gerado fica em **`dist\BuscaBranchGit_v3.exe`** (modo janela, sem console).

No **macOS/Linux** também é possível rodar o PyInstaller com o mesmo `.spec`, mas o hábito do projeto é empacotar o `.exe` no Windows.

## Notas

- O usuário final do `.exe` precisa ter **Git instalado** no PATH; o PyInstaller **não** embute o Git.
- Campos de pasta/arquivo na interface usam caminhos locais do clone; o padrão da pasta é `Documents` do usuário logado.
