# -*- coding: utf-8 -*-
"""
MathToWord — конвертер Markdown + LaTeX в Word с родными редактируемыми формулами (OMML).

Принцип работы (как у MassiveMark):
  1. Вставляете Markdown с формулами из нейросети.
  2. Нажимаете «⚡ Скопировать для Word» — форматированный текст с формулами попадает
     в буфер обмена, затем Ctrl+V в Word. Либо «💾 Сохранить в .docx» — файл на Рабочий стол.

Зависимости:
    pip install pypandoc pyinstaller
    (pandoc.exe должен лежать рядом с app.py для сборки)

Сборка Windows (один .exe):
    pyinstaller --noconsole --onefile --add-binary "pandoc.exe;." --name "MathToWord" app.py

Сборка macOS (.app) — через GitHub Actions, см. .github/workflows/build-macos.yml и MathToWord_mac.spec.
Копирование «для Word» на macOS работает через AppleScript (Microsoft Word для Mac).
"""

import os
import sys
import re
import time
import shutil
import tempfile
import threading
import queue
import subprocess
import ctypes
import traceback

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, scrolledtext

# --------------------------------------------------------------------------------------
# Защита для сборки с --noconsole: sys.stdout / sys.stderr в этом режиме равны None
# --------------------------------------------------------------------------------------
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

IS_WINDOWS = os.name == "nt"
IS_MAC = sys.platform == "darwin"
PANDOC_BINARY_NAME = "pandoc.exe" if IS_WINDOWS else "pandoc"
CMD_KEY_NAME = "Cmd" if IS_MAC else "Ctrl"
CREATE_NO_WINDOW = 0x08000000
APP_TITLE = "MathToWord — Markdown + LaTeX → Word (формулы OMML)"
TEMP_PREFIX = "MathToWord_"

# Все дочерние процессы (pandoc, powershell) запускаем без мигающего консольного окна.
if IS_WINDOWS:
    _OriginalPopen = subprocess.Popen

    class _HiddenWindowPopen(_OriginalPopen):
        def __init__(self, *args, **kwargs):
            kwargs["creationflags"] = kwargs.get("creationflags", 0) | CREATE_NO_WINDOW
            super().__init__(*args, **kwargs)

    subprocess.Popen = _HiddenWindowPopen


# --------------------------------------------------------------------------------------
# 1. АВТОНОМНОСТЬ: поиск вшитого pandoc.exe ДО импорта pypandoc
# --------------------------------------------------------------------------------------
def _find_bundled_pandoc():
    """Ищет pandoc: сначала в sys._MEIPASS (PyInstaller), затем рядом с exe/.app и скриптом."""
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, PANDOC_BINARY_NAME))
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        candidates.append(os.path.join(exe_dir, PANDOC_BINARY_NAME))
        if IS_MAC:  # MathToWord.app/Contents/MacOS -> ../Frameworks, ../Resources
            candidates.append(os.path.join(exe_dir, "..", "Frameworks", PANDOC_BINARY_NAME))
            candidates.append(os.path.join(exe_dir, "..", "Resources", PANDOC_BINARY_NAME))
    try:
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), PANDOC_BINARY_NAME))
    except NameError:
        pass
    for candidate in candidates:
        candidate = os.path.normpath(candidate)
        if os.path.isfile(candidate):
            return candidate
    return None


LOCAL_PANDOC_PATH = _find_bundled_pandoc()
if LOCAL_PANDOC_PATH:
    # Важно: переменная выставляется до первого обращения pypandoc к бинарнику.
    os.environ["PYPANDOC_PANDOC"] = LOCAL_PANDOC_PATH

try:
    import pypandoc
except Exception as _import_exc:  # pragma: no cover
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "MathToWord",
        "Не удалось импортировать pypandoc:\n%s\n\nУстановите: pip install pypandoc" % _import_exc,
    )
    sys.exit(1)

_pandoc_path_cache = None


def ensure_pandoc():
    """Возвращает путь к pandoc или None. Никогда ничего не скачивает."""
    global _pandoc_path_cache
    if _pandoc_path_cache:
        return _pandoc_path_cache
    if LOCAL_PANDOC_PATH:
        _pandoc_path_cache = LOCAL_PANDOC_PATH
        return _pandoc_path_cache
    try:
        path = pypandoc.get_pandoc_path()
    except OSError:
        path = None
    if not path:
        path = shutil.which("pandoc")
        if path:
            os.environ["PYPANDOC_PANDOC"] = path
    _pandoc_path_cache = path
    return path


class AppError(Exception):
    """Ошибка с понятным пользователю текстом."""


def task_download_pandoc():
    """Скачивание pandoc — ТОЛЬКО при запуске из исходников и ТОЛЬКО если pandoc нигде не найден."""
    global _pandoc_path_cache
    if LOCAL_PANDOC_PATH or getattr(sys, "frozen", False):
        raise AppError("Загрузка Pandoc запрещена: используется вшитый pandoc.exe.")
    pypandoc.download_pandoc()
    _pandoc_path_cache = None
    if not ensure_pandoc():
        raise AppError("Pandoc так и не удалось установить. Скачайте pandoc.exe вручную и положите рядом с app.py.")
    return True


# --------------------------------------------------------------------------------------
# 2. ПРЕДОБРАБОТКА: кириллица в формулах, склеенные команды, ограничители \( \) и \[ \]
# --------------------------------------------------------------------------------------
PH_OPEN, PH_CLOSE = "\ue000", "\ue001"      # «сейф» для \text{...}, \mathrm{...}
CODE_OPEN, CODE_CLOSE = "\ue002", "\ue003"  # защита блоков кода

# Все символы кириллицы (включая украинскую, казахскую и т.д.)
_CYR = "\u0400-\u052f"
_CYR_RE = re.compile("[" + _CYR + "]")
# «Голая» последовательность русских букв; внутри допускаются точка и пробел между словами (ош.п, тр.зс)
_CYR_RUN_RE = re.compile("[" + _CYR + "]+(?:[ \u00a0.][" + _CYR + "]+)*")

# Теги, содержимое которых прячется в сейф (чтобы не получить \text{\text{...}})
_PROTECTED_TAGS = (
    "text", "textrm", "textbf", "textit", "textsf", "texttt", "textnormal", "mbox",
    "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "operatorname",
)
_TAG_OPEN_RE = re.compile(
    r"\\(" + "|".join(sorted(_PROTECTED_TAGS, key=len, reverse=True)) + r")(?![A-Za-z])\s*\{"
)
# При возврате из сейфа математические шрифты с кириллицей превращаются в текстовые аналоги,
# иначе Pandoc снова стирает русские буквы.
_MATH_FONT_TO_TEXT = {
    "mathrm": "text", "mathbf": "textbf", "mathit": "textit",
    "mathsf": "textsf", "mathtt": "texttt", "operatorname": "text",
}
_PH_RE = re.compile(PH_OPEN + r"(\d+)" + PH_CLOSE)
_CODE_PH_RE = re.compile(CODE_OPEN + r"(\d+)" + CODE_CLOSE)

# Команды, которые браузер/нейросеть часто «склеивает» со следующими буквами (\cdotlg, \Deltaf)
SPLIT_COMMANDS = frozenset("""
alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi pi varpi
rho varrho sigma varsigma tau upsilon phi varphi chi psi omega
Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega
cdot times div pm mp ast star circ bullet oplus otimes cap cup
leq geq neq approx equiv sim simeq propto ll gg le ge ne to in notin subset supset
sin cos tan cot sec csc arcsin arccos arctan sinh cosh tanh coth log ln lg lim exp max min sup inf
det dim ker deg gcd arg
sum prod int oint partial nabla infty sqrt frac dfrac tfrac
ldots cdots dots rightarrow leftarrow Rightarrow Leftarrow leftrightarrow Leftrightarrow
hbar ell
""".split())

# Все известные команды (чтобы не «разрезать» корректные команды вроде \intercal или \leftarrow)
KNOWN_COMMANDS = SPLIT_COMMANDS | frozenset("""
left right big Big bigg Bigg bigl bigr Bigl Bigr biggl biggr Biggl Biggr middle
text textrm textbf textit textsf texttt textnormal mbox hbox fbox framebox
mathrm mathbf mathit mathsf mathtt mathcal mathbb mathfrak mathscr boldsymbol bm pmb
operatorname overline underline overbrace underbrace overleftarrow overrightarrow
widehat widetilde hat tilde bar vec dot ddot dddot acute grave breve check mathring
begin end limits nolimits displaystyle textstyle scriptstyle scriptscriptstyle
binom dbinom tbinom choose atop over brace brack cfrac root
quad qquad hspace hskip mspace thinspace medspace thickspace negthinspace enspace space
phantom vphantom hphantom smash mathstrut strut
xrightarrow xleftarrow substack sideset stackrel overset underset boxed cancel bcancel xcancel
not pmod bmod mod pmatrix tag label nonumber notag intertext eqno leqno lefteqn
lbrace rbrace lbrack rbrack langle rangle lvert rvert lVert rVert lfloor rfloor lceil rceil
vert Vert backslash setminus emptyset varnothing forall exists nexists neg lnot land lor wedge vee
implies iff Longrightarrow Longleftarrow longrightarrow longleftarrow Longleftrightarrow longleftrightarrow
mapsto longmapsto uparrow downarrow Uparrow Downarrow updownarrow Updownarrow
nearrow searrow swarrow nwarrow hookrightarrow hookleftarrow rightarrowtail leftarrowtail
leftharpoondown leftharpoonup rightharpoondown rightharpoonup leftleftarrows leftrightarrows
leftrightharpoons rightleftharpoons leftrightsquigarrow leftroot leftthreetimes
leqslant geqslant leqq geqq lll ggg ggg llap llbracket llcorner
lessapprox lessdot lesseqgtr lesseqqgtr lessgtr lesssim gtrsim gtrless gtrdot
prec succ preceq succeq cong ncong nsim nequiv asymp doteq approxeq thickapprox thicksim
coloneqq eqqcolon triangle triangleq angle measuredangle perp parallel nparallel mid nmid
subseteq supseteq subsetneq supsetneq subseteqq supseteqq subsetneqq supsetneqq nsubseteq nsupseteq
sqsubset sqsupset sqcap sqcup uplus bigcup bigcap bigsqcup bigvee bigwedge bigoplus bigotimes bigodot biguplus
iint iiint iiiint idotsint oiint oiiint intercal intop injlim top bot vdash dashv models vDash Vdash
multimap multicolumn multirow pitchfork numberwithin genfrac gets newline newcommand renewcommand let
nearrow circeq circlearrowleft circlearrowright circledast circledcirc circleddash circledR circledS
cdotp cdots centerdot dotsc dotsb dotsm dotsi dotso vdots ddots mathellipsis colon ldotp surd
lgroup rgroup lneq lneqq lnsim lnapprox limsup liminf varlimsup varliminf varinjlim varprojlim projlim
Pr argmax argmin sgn tg ctg arctg arcctg lb th sh ch cth
aleph beth gimel daleth Re Im wp prime backprime hslash imath jmath dag ddag
copyright pounds degree celsius textdegree checkmark sharp flat natural
clubsuit diamondsuit heartsuit spadesuit bigstar therefore because
color textcolor colorbox rule raisebox item hline cline noindent par ensuremath unicode char
sec csc Bbbk
""".split())

_CMD_OR_ESCAPE_RE = re.compile(r"(\\\\)|\\([A-Za-z]+)")


def _find_closing_brace(s, open_idx):
    """Индекс парной закрывающей скобки для s[open_idx] == '{' (с учётом вложенности и \\{ \\})."""
    depth = 0
    i = open_idx
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _vault_text_tags(s, vault):
    """Шаг (а): прячем существующие \\text{...}, \\mathrm{...} в сейф, заменяя плейсхолдерами."""
    out = []
    pos = 0
    while True:
        m = _TAG_OPEN_RE.search(s, pos)
        if not m:
            out.append(s[pos:])
            break
        open_idx = m.end() - 1
        close_idx = _find_closing_brace(s, open_idx)
        if close_idx == -1:
            # Несбалансированные скобки — оставляем как есть
            out.append(s[pos:m.end()])
            pos = m.end()
            continue
        out.append(s[pos:m.start()])
        vault.append((m.group(1), s[open_idx + 1:close_idx]))
        out.append("%s%d%s" % (PH_OPEN, len(vault) - 1, PH_CLOSE))
        pos = close_idx + 1
    return "".join(out)


def _wrap_cyrillic(s):
    """Шаг (б): оборачиваем в \\text{...} только оставшиеся «голые» русские последовательности."""
    return _CYR_RUN_RE.sub(lambda m: "\\text{" + m.group(0) + "}", s)


def _split_rest(rest):
    parts = []
    while rest:
        cut = 0
        for i in range(len(rest), 1, -1):
            if rest[:i] in SPLIT_COMMANDS:
                cut = i
                break
        if cut:
            parts.append(("cmd", rest[:cut]))
            rest = rest[cut:]
        else:
            parts.append(("txt", rest))
            break
    return parts


def _split_glued(name):
    for i in range(len(name) - 1, 1, -1):
        if name[:i] in SPLIT_COMMANDS:
            return [("cmd", name[:i])] + _split_rest(name[i:])
    return None


def _split_glued_commands(s):
    """Шаг (в): \\cdotlg -> \\cdot \\lg, \\Deltaf -> \\Delta f."""
    def repl(m):
        if m.group(1):          # \\ (перенос строки) — не трогаем
            return m.group(0)
        name = m.group(2)
        if name in KNOWN_COMMANDS:
            return m.group(0)
        parts = _split_glued(name)
        if not parts:
            return m.group(0)
        return " ".join(("\\" + value) if kind == "cmd" else value for kind, value in parts)

    return _CMD_OR_ESCAPE_RE.sub(repl, s)


def _unvault_text_tags(s, vault):
    """Шаг (г): возвращаем спрятанные теги из плейсхолдеров."""
    def repl(m):
        cmd, inner = vault[int(m.group(1))]
        if cmd in _MATH_FONT_TO_TEXT and _CYR_RE.search(inner) and "\\" not in inner:
            cmd = _MATH_FONT_TO_TEXT[cmd]
        return "\\" + cmd + "{" + inner + "}"

    return _PH_RE.sub(repl, s)


def _process_math(body):
    vault = []
    body = _vault_text_tags(body, vault)   # а) сейф
    body = _wrap_cyrillic(body)            # б) \text{...} для голой кириллицы
    body = _split_glued_commands(body)     # в) разделение слипшихся команд
    body = _unvault_text_tags(body, vault) # г) возврат из сейфа
    return body


_FENCED_CODE_RE = re.compile(r"^[ \t]*(`{3,}|~{3,})[^\n]*\n.*?^[ \t]*\1[ \t]*$", re.M | re.S)
_INLINE_CODE_RE = re.compile(r"(`+)(?!`)[^\n]*?(?<!`)\1(?!`)")
_BS_DISPLAY_RE = re.compile(r"(?<!\\)\\\[(.+?)(?<!\\)\\\]", re.S)
_BS_INLINE_RE = re.compile(r"(?<!\\)\\\((.+?)(?<!\\)\\\)", re.S)
_MATH_SPAN_RE = re.compile(
    r"(?<!\\)\$\$((?:\\[\s\S]|[^\\$]|\$(?!\$))+?)\$\$"
    r"|(?<![\\$])\$(?![\s$])((?:\\[\s\S]|[^$\\\n]|\n(?![ \t]*\n))+?)(?<!\s)\$(?!\d)"
)


def _math_span_callback(m):
    if m.group(1) is not None:
        return "$$" + _process_math(m.group(1)) + "$$"
    return "$" + _process_math(m.group(2)) + "$"


def preprocess_markdown(text):
    """Полная предобработка Markdown: возвращает «вылеченный» текст для Pandoc."""
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\ufeff", "")
    for ch in (PH_OPEN, PH_CLOSE, CODE_OPEN, CODE_CLOSE):
        text = text.replace(ch, "")

    # Блоки кода не должны затрагиваться (в них могут быть $ и кириллица)
    code_store = []

    def stash(m):
        code_store.append(m.group(0))
        return "%s%d%s" % (CODE_OPEN, len(code_store) - 1, CODE_CLOSE)

    text = _FENCED_CODE_RE.sub(stash, text)
    text = _INLINE_CODE_RE.sub(stash, text)

    # \[ ... \] -> $$ ... $$,  \( ... \) -> $ ... $
    text = _BS_DISPLAY_RE.sub(lambda m: "$$" + m.group(1).strip() + "$$", text)
    text = _BS_INLINE_RE.sub(lambda m: "$" + m.group(1).strip() + "$", text)

    text = _MATH_SPAN_RE.sub(_math_span_callback, text)

    text = _CODE_PH_RE.sub(lambda m: code_store[int(m.group(1))], text)
    return text


# --------------------------------------------------------------------------------------
# Утилиты: имя файла, рабочий стол, блокировка файла
# --------------------------------------------------------------------------------------
_FORBIDDEN_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {"COM%d" % i for i in range(1, 10)} | {"LPT%d" % i for i in range(1, 10)}


def sanitize_filename(raw):
    """Очищает имя от запрещённых символов и гарантированно добавляет расширение .docx."""
    name = _FORBIDDEN_RE.sub("", raw or "").strip().rstrip(". ")
    while name.lower().endswith(".docx"):
        name = name[:-5].rstrip(". ")
    name = name.strip()
    if not name:
        name = "Документ"
    if name.split(".")[0].upper() in _RESERVED_NAMES:
        name = "_" + name
    return name[:150].rstrip(". ") + ".docx"


def get_desktop_dir():
    """Физическая папка Рабочего стола (учитывает перенаправление в OneDrive и локализацию)."""
    if IS_WINDOWS:
        try:
            buf = ctypes.create_unicode_buffer(1024)
            # CSIDL_DESKTOPDIRECTORY = 0x0010
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0:
                if buf.value and os.path.isdir(buf.value):
                    return buf.value
        except Exception:
            pass
    home = os.path.expanduser("~")
    for candidate in (
        os.path.join(home, "Desktop"),
        os.path.join(home, "Рабочий стол"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive", "Рабочий стол"),
    ):
        if os.path.isdir(candidate):
            return candidate
    return home


def is_file_locked(path):
    """True, если файл существует и занят другим процессом (например, открыт в Word)."""
    if not os.path.exists(path):
        return False
    if IS_MAC:
        # Word для Mac не блокирует файл на уровне ОС, но создаёт рядом служебный файл «~$...»
        folder, name = os.path.split(path)
        for owner in ("~$" + name[2:], "~$" + name):
            if os.path.exists(os.path.join(folder, owner)):
                return True
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def safe_rmtree(path):
    for _ in range(6):
        try:
            shutil.rmtree(path)
            return
        except FileNotFoundError:
            return
        except OSError:
            time.sleep(0.3)
    shutil.rmtree(path, ignore_errors=True)


def cleanup_stale_temp_dirs(max_age_seconds=6 * 3600):
    """Удаляет забытые рабочие папки прошлых аварийных запусков."""
    try:
        base = tempfile.gettempdir()
        now = time.time()
        for entry in os.listdir(base):
            if not entry.startswith(TEMP_PREFIX):
                continue
            full = os.path.join(base, entry)
            try:
                if os.path.isdir(full) and now - os.path.getmtime(full) > max_age_seconds:
                    shutil.rmtree(full, ignore_errors=True)
            except OSError:
                pass
    except Exception:
        pass


# --------------------------------------------------------------------------------------
# 3. КОНВЕРТАЦИЯ ЧЕРЕЗ ФАЙЛЫ НА ДИСКЕ + КОПИРОВАНИЕ В БУФЕР WORD
# --------------------------------------------------------------------------------------
# Сначала пробуем отключить yaml_metadata_block: строки «---» из ответов нейросетей
# иначе могут быть приняты за YAML-блок. Если Pandoc не примет расширения — запасной вариант.
PANDOC_INPUT_FORMATS = ("markdown-yaml_metadata_block+tex_math_dollars", "markdown")


def convert_markdown_to_docx(md_text, out_path, work_dir):
    """Markdown -> docx через pypandoc.convert_file (файл на диске, а не аргументы командной строки)."""
    if not ensure_pandoc():
        raise AppError(
            "Pandoc не найден. Положите %s рядом с app.py (для готовой сборки он вшивается "
            "при компиляции: см. инструкцию по сборке в начале app.py)." % PANDOC_BINARY_NAME
        )
    md_path = os.path.join(work_dir, "source.md")
    with open(md_path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(md_text)

    last_error = None
    for fmt in PANDOC_INPUT_FORMATS:
        try:
            pypandoc.convert_file(md_path, "docx", format=fmt, outputfile=out_path)
            last_error = None
            break
        except RuntimeError as exc:
            last_error = exc
        except OSError as exc:
            raise AppError("Не удалось запустить Pandoc: %s" % exc)
    if last_error is not None:
        raise AppError("Pandoc не смог выполнить конвертацию:\n%s" % str(last_error)[:1200])
    if not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
        raise AppError("Pandoc завершился без ошибки, но файл .docx не был создан.")


# PowerShell-скрипт: перехват запущенного Word или фоновый запуск, Content.Copy(),
# пауза 300 мс против Delayed Rendering, закрытие документа без сохранения.
PS_COPY_SCRIPT = r"""
param(
    [Parameter(Mandatory = $true)]
    [string]$DocPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$word = $null
$doc = $null
$createdWord = $false
$copied = $false
$stage = 'start'
$exitCode = 0

try {
    # 1. Try to attach to an already running Word instance
    $stage = 'attach'
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    } catch {
        $word = $null
    }

    # 2. Word is closed - create a hidden background instance
    if ($null -eq $word) {
        $stage = 'create'
        $word = New-Object -ComObject Word.Application
        $createdWord = $true
        $word.Visible = $false
        try { $word.DisplayAlerts = 0 } catch { }
    }

    # 3. Open the temporary document (read-only, hidden window, not added to recent files)
    $stage = 'open'
    $missing = [System.Reflection.Missing]::Value
    try {
        $doc = $word.Documents.Open($DocPath, $false, $true, $false, $missing, $missing, $false, $missing, $missing, $missing, $missing, $false)
    } catch {
        $doc = $word.Documents.Open($DocPath)
    }

    # 4. Copy the whole content to the clipboard
    $stage = 'copy'
    $doc.Content.Copy()
    $copied = $true

    # 5. CRITICAL: work around the Windows "Delayed Rendering" bug.
    #    Without this pause the clipboard is emptied when the document is closed.
    Start-Sleep -Milliseconds 300
}
catch {
    [Console]::Error.WriteLine('STAGE=' + $stage + '; ' + $_.Exception.Message)
    if ($stage -eq 'create') { $exitCode = 10 }
    elseif ($stage -eq 'open') { $exitCode = 11 }
    elseif ($stage -eq 'copy') { $exitCode = 12 }
    else { $exitCode = 1 }
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close(0) } catch { try { $doc.Close([ref]0) } catch { } }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
    }
    if ($createdWord -and $null -ne $word) {
        Start-Sleep -Milliseconds 300
        try { $word.Quit(0) } catch { try { $word.Quit([ref]0) } catch { } }
    }
    if ($null -ne $word) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    }
}
exit $exitCode
"""

_PS_ERROR_TEXTS = {
    10: "Microsoft Word не найден или не запускается (COM-объект Word.Application недоступен). "
        "Установите Word либо используйте кнопку «Сохранить в .docx».",
    11: "Word не смог открыть временный документ. Закройте диалоговые окна в Word и повторите попытку.",
    12: "Word не смог скопировать содержимое в буфер обмена. Повторите попытку.",
}


def _decode_console_bytes(data):
    if not data:
        return ""
    try:
        return data.decode("utf-8").lstrip("\ufeff").strip()
    except UnicodeDecodeError:
        try:
            oem = ctypes.windll.kernel32.GetOEMCP()
            return data.decode("cp%d" % oem, errors="replace").strip()
        except Exception:
            return data.decode("utf-8", errors="replace").strip()


def copy_docx_to_clipboard_via_word(docx_path, work_dir):
    """Открывает временный .docx в Word (COM) и копирует его содержимое в буфер обмена."""
    script_path = os.path.join(work_dir, "copy_to_clipboard.ps1")
    # utf-8-sig: Windows PowerShell 5.1 корректно читает скрипт только с BOM
    with open(script_path, "w", encoding="utf-8-sig", newline="\r\n") as handle:
        handle.write(PS_COPY_SCRIPT)

    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    powershell = os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    if not os.path.isfile(powershell):
        powershell = "powershell"

    command = [
        powershell, "-NoProfile", "-NonInteractive", "-STA",
        "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
        "-File", script_path, "-DocPath", docx_path,
    ]
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0  # SW_HIDE

    try:
        proc = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=180,
            startupinfo=startupinfo,
            creationflags=CREATE_NO_WINDOW,
        )
    except FileNotFoundError:
        raise AppError("Не найден PowerShell (powershell.exe). Он необходим для копирования в буфер Word.")
    except subprocess.TimeoutExpired:
        raise AppError("Word не ответил за 3 минуты. Закройте диалоговые окна в Word и повторите попытку.")

    if proc.returncode != 0:
        details = _decode_console_bytes(proc.stderr) or _decode_console_bytes(proc.stdout)
        message = _PS_ERROR_TEXTS.get(proc.returncode, "Не удалось скопировать в буфер обмена через Word.")
        if details:
            message += "\n\nПодробности: " + details[:600]
        raise AppError(message)


# AppleScript для macOS: открыть временный .docx в Microsoft Word и скопировать всё содержимое.
# Основной способ — команда Word «copy object»; запасной — Cmd+A / Cmd+C через System Events
# (по физическим key code 0 и 8, поэтому работает на любой раскладке).
AS_COPY_SCRIPT = r"""
on run argv
    set posixPath to item 1 of argv
    set theFile to POSIX file posixPath

    tell application "Microsoft Word"
        activate
        set beforeCount to count of documents
        open theFile
        set waited to 0
        repeat while ((count of documents) <= beforeCount) and (waited < 150)
            delay 0.1
            set waited to waited + 1
        end repeat
        if (count of documents) <= beforeCount then error "STAGE=open" number 11
        set theDoc to active document
    end tell

    set copied to false
    try
        tell application "Microsoft Word"
            copy object (text object of theDoc)
        end tell
        set copied to true
    end try

    if not copied then
        try
            tell application "Microsoft Word" to activate
            delay 0.5
            tell application "System Events"
                tell process "Microsoft Word"
                    set frontmost to true
                    key code 0 using {command down}
                    delay 0.2
                    key code 8 using {command down}
                end tell
            end tell
            set copied to true
        on error errMsg number errNum
            error "STAGE=copy; " & errMsg & " (" & errNum & ")" number 12
        end try
    end if

    -- Pause so the pasteboard data is fully written before the document is closed
    delay 0.5
    tell application "Microsoft Word"
        try
            close theDoc saving no
        end try
    end tell
    return "OK"
end run
"""

_AS_ERROR_TEXTS = {
    11: "Word не смог открыть временный документ. Закройте диалоговые окна в Word и повторите попытку.",
    12: "Не удалось скопировать содержимое через Word. Если macOS запросила разрешение — разрешите его "
        "(Системные настройки → Конфиденциальность и безопасность → Универсальный доступ / Автоматизация).",
}


def _mac_word_installed():
    try:
        proc = subprocess.run(
            ["osascript", "-e", 'id of application "Microsoft Word"'],
            stdin=subprocess.DEVNULL, capture_output=True, timeout=30,
        )
        return proc.returncode == 0
    except Exception:
        return False


def copy_docx_to_clipboard_via_word_mac(docx_path, work_dir):
    """macOS: открывает временный .docx в Word через AppleScript и копирует содержимое в буфер."""
    if not _mac_word_installed():
        raise AppError(
            "Microsoft Word для Mac не найден. Установите Word либо используйте кнопку «Сохранить в .docx»."
        )
    script_path = os.path.join(work_dir, "copy_to_clipboard.applescript")
    with open(script_path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(AS_COPY_SCRIPT)
    try:
        proc = subprocess.run(
            ["osascript", script_path, docx_path],
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
        )
    except FileNotFoundError:
        raise AppError("Не найден osascript — копирование через Word невозможно.")
    except subprocess.TimeoutExpired:
        raise AppError("Word не ответил за 3 минуты. Закройте диалоговые окна в Word и повторите попытку.")

    if proc.returncode != 0:
        details = proc.stderr.decode("utf-8", errors="replace").strip()
        if "-1743" in details or "not authorized" in details.lower():
            raise AppError(
                "macOS не разрешила MathToWord управлять Microsoft Word.\n\n"
                "Откройте Системные настройки → Конфиденциальность и безопасность → Автоматизация "
                "и включите Microsoft Word для MathToWord (или для Терминала при запуске из исходников). "
                "Затем повторите попытку."
            )
        message = "Не удалось скопировать в буфер обмена через Word."
        for code, text in _AS_ERROR_TEXTS.items():
            if "(%d)" % code in details or "number %d" % code in details:
                message = text
                break
        if details:
            message += "\n\nПодробности: " + details[:600]
        raise AppError(message)


def open_with_default_app(path):
    """Открывает файл программой по умолчанию (Word) на любой платформе."""
    if IS_WINDOWS:
        os.startfile(path)
    elif IS_MAC:
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# --------------------------------------------------------------------------------------
# Фоновые задачи (выполняются в отдельном потоке)
# --------------------------------------------------------------------------------------
def task_copy_for_word(md_text):
    if not (IS_WINDOWS or IS_MAC):
        raise AppError("Копирование в буфер Word работает только в Windows и macOS с установленным Microsoft Word.")
    work_dir = tempfile.mkdtemp(prefix=TEMP_PREFIX)
    try:
        cured = preprocess_markdown(md_text)
        docx_path = os.path.join(work_dir, "clipboard_source.docx")
        convert_markdown_to_docx(cured, docx_path, work_dir)
        if IS_MAC:
            copy_docx_to_clipboard_via_word_mac(docx_path, work_dir)
        else:
            copy_docx_to_clipboard_via_word(docx_path, work_dir)
        return len(md_text)
    finally:
        safe_rmtree(work_dir)  # временные файлы из Temp удаляются всегда


def task_save_docx(md_text, dest_path):
    work_dir = tempfile.mkdtemp(prefix=TEMP_PREFIX)
    try:
        cured = preprocess_markdown(md_text)
        tmp_docx = os.path.join(work_dir, "result.docx")
        convert_markdown_to_docx(cured, tmp_docx, work_dir)
        try:
            shutil.copyfile(tmp_docx, dest_path)
        except PermissionError:
            raise AppError(
                "Не удалось записать файл «%s».\nОн открыт в Word (или защищён от записи). "
                "Закройте документ и повторите сохранение." % os.path.basename(dest_path)
            )
        except OSError as exc:
            raise AppError("Не удалось сохранить файл: %s" % exc)
        return dest_path
    finally:
        safe_rmtree(work_dir)


def task_cure_markdown(md_text):
    return preprocess_markdown(md_text)


# --------------------------------------------------------------------------------------
# 6. ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# --------------------------------------------------------------------------------------
class MathToWordApp:
    STATUS_COLORS = {"info": "#333333", "ok": "#1b7f3b", "warn": "#a15c00", "error": "#b3261e"}

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1040x720")
        self.root.minsize(800, 460)

        self._queue = queue.Queue()
        self._busy = False
        self._closing = False
        self.name_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Готово. Вставьте Markdown и выберите действие.")

        self._setup_style()
        self.mono_font = self._pick_mono_font()
        self._build_ui()
        self._build_context_menu()
        self._bind_hotkeys()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(80, self._poll_queue)
        self.root.after(250, self._startup_check)
        threading.Thread(target=cleanup_stale_temp_dirs, daemon=True).start()

    # ---------------------------------------------------------------- построение окна
    def _setup_style(self):
        style = ttk.Style(self.root)
        if IS_WINDOWS:
            for theme in ("vista", "winnative", "clam"):
                if theme in style.theme_names():
                    try:
                        style.theme_use(theme)
                        break
                    except tk.TclError:
                        continue
        # На macOS остаётся родная тема aqua.
        self._bold_font = tkfont.nametofont("TkDefaultFont").copy()
        self._bold_font.configure(weight="bold")
        if not IS_MAC:
            style.configure("TButton", padding=(10, 6))
        style.configure("Accent.TButton", font=self._bold_font)
        if not IS_MAC:
            style.configure("Accent.TButton", padding=(12, 6))

    def _pick_mono_font(self):
        available = set(tkfont.families(self.root))
        for family in ("Consolas", "Menlo", "SF Mono", "Cascadia Mono", "Cascadia Code", "Monaco", "Courier New", "DejaVu Sans Mono"):
            if family in available:
                return tkfont.Font(root=self.root, family=family, size=11)
        fallback = tkfont.nametofont("TkFixedFont").copy()
        fallback.configure(size=11)
        return fallback

    def _build_ui(self):
        # --- верхняя панель
        top = ttk.Frame(self.root, padding=(10, 10, 10, 6))
        top.pack(fill="x")
        ttk.Label(top, text="Имя файла:").pack(side="left")
        self.btn_paste = ttk.Button(top, text="Вставить исходник", command=self.on_paste_source)
        self.btn_paste.pack(side="right")
        ttk.Label(top, text=".docx добавится автоматически", foreground="#777777").pack(side="right", padx=(8, 14))
        vcmd = (self.root.register(self._validate_name), "%P")
        self.name_entry = ttk.Entry(top, textvariable=self.name_var, validate="key", validatecommand=vcmd)
        self.name_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.name_entry.bind("<FocusOut>", self._normalize_name_entry)

        # --- центральное поле ввода (undo отключён, чтобы не копить историю на гигантских текстах)
        self.text = scrolledtext.ScrolledText(
            self.root, undo=False, wrap="word", font=self.mono_font,
            padx=8, pady=6, borderwidth=1, relief="solid",
        )
        self.text.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # --- нижняя панель
        bottom = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        bottom.pack(fill="x")
        self.btn_copy_text = ttk.Button(bottom, text="📝 Скопировать текст", command=self.on_copy_text)
        self.btn_copy_text.pack(side="right")
        self.btn_save = ttk.Button(bottom, text="💾 Сохранить в .docx", command=self.on_save)
        self.btn_save.pack(side="right", padx=(0, 8))
        self.btn_copy_word = ttk.Button(
            bottom, text="⚡ Скопировать для Word", style="Accent.TButton", command=self.on_copy_for_word
        )
        self.btn_copy_word.pack(side="right", padx=(0, 8))
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=110)
        self.status_label = ttk.Label(bottom, textvariable=self.status_var, anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)
        self._action_buttons = (self.btn_paste, self.btn_copy_word, self.btn_save, self.btn_copy_text)

    def _build_context_menu(self):
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Вырезать", command=lambda: self._hk_cut(self.text))
        self.menu.add_command(label="Копировать", command=lambda: self._hk_copy(self.text))
        self.menu.add_command(label="Вставить", command=lambda: self._hk_paste(self.text))
        self.menu.add_separator()
        self.menu.add_command(label="Выделить всё", command=lambda: self._hk_select_all(self.text))
        if IS_MAC:
            self.text.bind("<Button-2>", self._show_context_menu)
            self.text.bind("<Control-Button-1>", self._show_context_menu)
        else:
            self.text.bind("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event):
        self.text.focus_set()
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    # ---------------------------------------------------------------- горячие клавиши
    def _bind_hotkeys(self):
        """Windows: Ctrl+A/V/C/X по физическим кодам клавиш (работает на любой раскладке).
        macOS: Cmd+A/V/C/X, в том числе на русской раскладке (по символу и keysym)."""
        if IS_WINDOWS:
            for widget in (self.text, self.name_entry):
                widget.bind("<Control-KeyPress>", self._on_ctrl_key)
        elif IS_MAC:
            for widget in (self.text, self.name_entry):
                widget.bind("<Command-KeyPress>", self._on_cmd_key_mac)

    def _on_ctrl_key(self, event):
        if event.state & 0x20000:   # Alt / AltGr — не перехватываем
            return None
        handlers = {
            65: self._hk_select_all,   # A
            86: self._hk_paste,        # V
            67: self._hk_copy,         # C
            88: self._hk_cut,          # X
        }
        handler = handlers.get(event.keycode)
        if handler is None:
            return None
        try:
            handler(event.widget)
        except tk.TclError:
            pass
        return "break"

    # Положение клавиш ЙЦУКЕН: ф=A, с=C, м=V, ч=X
    _MAC_KEY_MAP = {
        "a": "a", "ф": "a", "cyrillic_ef": "a",
        "c": "c", "с": "c", "cyrillic_es": "c",
        "v": "v", "м": "v", "cyrillic_em": "v",
        "x": "x", "ч": "x", "cyrillic_che": "x",
    }

    def _on_cmd_key_mac(self, event):
        key = self._MAC_KEY_MAP.get((event.keysym or "").lower()) or self._MAC_KEY_MAP.get((event.char or "").lower())
        handlers = {
            "a": self._hk_select_all,
            "v": self._hk_paste,
            "c": self._hk_copy,
            "x": self._hk_cut,
        }
        handler = handlers.get(key)
        if handler is None:
            return None
        try:
            handler(event.widget)
        except tk.TclError:
            pass
        return "break"

    def _hk_select_all(self, widget):
        if isinstance(widget, tk.Text):
            widget.tag_add("sel", "1.0", "end-1c")
            widget.mark_set("insert", "end-1c")
        else:
            widget.select_range(0, "end")
            widget.icursor("end")

    @staticmethod
    def _selected_text(widget):
        if isinstance(widget, tk.Text):
            if widget.tag_ranges("sel"):
                return widget.get("sel.first", "sel.last")
            return ""
        if widget.selection_present():
            start, end = widget.index("sel.first"), widget.index("sel.last")
            return widget.get()[start:end]
        return ""

    def _hk_copy(self, widget):
        selected = self._selected_text(widget)
        if selected:
            self._clipboard_set(selected)

    def _hk_cut(self, widget):
        selected = self._selected_text(widget)
        if not selected:
            return
        self._clipboard_set(selected)
        if isinstance(widget, tk.Text):
            widget.delete("sel.first", "sel.last")
        else:
            widget.delete("sel.first", "sel.last")

    def _hk_paste(self, widget):
        try:
            data = self.root.clipboard_get()
        except tk.TclError:
            self._set_status("Буфер обмена пуст или содержит не текст.", "warn")
            return
        if isinstance(widget, tk.Text):
            data = data.replace("\r\n", "\n").replace("\r", "\n")
            if widget.tag_ranges("sel"):
                widget.delete("sel.first", "sel.last")
            widget.insert("insert", data)
            widget.see("insert")
        else:
            stripped = data.strip()
            first_line = stripped.splitlines()[0] if stripped else ""
            cleaned = _FORBIDDEN_RE.sub("", first_line)
            if widget.selection_present():
                widget.delete("sel.first", "sel.last")
            widget.insert("insert", cleaned)

    def _clipboard_set(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()

    # ---------------------------------------------------------------- имя файла
    def _validate_name(self, proposed):
        if _FORBIDDEN_RE.search(proposed):
            self.root.bell()
            return False
        return True

    def _normalize_name_entry(self, _event=None):
        if self.name_var.get().strip():
            self.name_var.set(sanitize_filename(self.name_var.get()))

    # ---------------------------------------------------------------- статус / занятость
    def _set_status(self, text, kind="info"):
        self.status_var.set(text)
        self.status_label.configure(foreground=self.STATUS_COLORS.get(kind, self.STATUS_COLORS["info"]))

    def _set_busy(self, busy, text=None):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for button in self._action_buttons:
            button.configure(state=state)
        if busy:
            self.progress.pack(side="left", padx=(0, 10))
            self.progress.start(12)
            if text:
                self._set_status(text, "info")
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def _get_text(self):
        return self.text.get("1.0", "end-1c")

    # ---------------------------------------------------------------- потоки
    def _start_job(self, busy_text, func, on_success, *args):
        if self._busy:
            return

        self._set_busy(True, busy_text)

        def runner():
            try:
                result = func(*args)
            except Exception as exc:
                self._queue.put(("error", exc, traceback.format_exc()))
            else:
                self._queue.put(("ok", on_success, result))

        threading.Thread(target=runner, name="MathToWordWorker", daemon=True).start()

    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                self._set_busy(False)
                if item[0] == "ok":
                    try:
                        item[1](item[2])
                    except Exception as exc:
                        self._handle_error(exc, traceback.format_exc())
                else:
                    self._handle_error(item[1], item[2])
        except queue.Empty:
            pass
        finally:
            if not self._closing:
                self.root.after(80, self._poll_queue)

    def _handle_error(self, exc, _traceback_text):
        if isinstance(exc, AppError):
            message = str(exc)
        else:
            message = "Непредвиденная ошибка: %s: %s" % (type(exc).__name__, exc)
        short = message.strip().splitlines()[0] if message.strip() else "Ошибка"
        self._set_status("❌ " + short[:140], "error")
        messagebox.showerror("MathToWord", message[:1500])

    # ---------------------------------------------------------------- старт
    def _startup_check(self):
        if ensure_pandoc():
            return
        if getattr(sys, "frozen", False):
            self._set_status("❌ Встроенный %s не найден." % PANDOC_BINARY_NAME, "error")
            messagebox.showerror(
                "MathToWord",
                "Встроенный %s не найден внутри приложения.\n\n"
                "Пересоберите приложение, включив pandoc в сборку (см. инструкцию в начале app.py)."
                % PANDOC_BINARY_NAME,
            )
            return
        if messagebox.askyesno(
            "Pandoc не найден",
            "Pandoc не найден ни рядом с app.py, ни в PATH.\n\n"
            "Скачать и установить его сейчас? (Либо положите %s рядом с app.py.)" % PANDOC_BINARY_NAME,
        ):
            self._start_job("⏳ Загрузка Pandoc…", task_download_pandoc, self._after_download_pandoc)
        else:
            self._set_status("⚠ Pandoc не найден — конвертация недоступна.", "warn")

    def _after_download_pandoc(self, _result):
        self._set_status("✅ Pandoc установлен. Можно работать.", "ok")

    # ---------------------------------------------------------------- действия
    def on_paste_source(self):
        try:
            data = self.root.clipboard_get()
        except tk.TclError:
            self._set_status("Буфер обмена пуст или содержит не текст.", "warn")
            return
        if not data.strip():
            self._set_status("Буфер обмена пуст.", "warn")
            return
        data = data.replace("\r\n", "\n").replace("\r", "\n")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", data)
        self.text.mark_set("insert", "1.0")
        self.text.see("1.0")
        self.text.focus_set()
        self._set_status("Вставлено %s символов." % format(len(data), ",").replace(",", " "), "ok")

    def on_copy_for_word(self):
        md_text = self._get_text()
        if not md_text.strip():
            self._set_status("Поле пустое — вставьте Markdown.", "warn")
            return
        self._start_job(
            "⏳ Конвертация и копирование в буфер Word…",
            task_copy_for_word, self._after_copy_for_word, md_text,
        )

    def _after_copy_for_word(self, _result):
        self._set_status("✅ Скопировано! Откройте документ Word и нажмите %s+V." % CMD_KEY_NAME, "ok")

    def on_save(self):
        md_text = self._get_text()
        if not md_text.strip():
            self._set_status("Поле пустое — вставьте Markdown.", "warn")
            return
        file_name = sanitize_filename(self.name_var.get())
        self.name_var.set(file_name)
        dest = os.path.join(get_desktop_dir(), file_name)

        if os.path.exists(dest):
            if is_file_locked(dest):
                self._set_status("⚠ Файл «%s» открыт в Word." % file_name, "warn")
                messagebox.showwarning(
                    "Файл открыт в Word",
                    "Файл «%s» сейчас открыт в Word (или занят другой программой).\n\n"
                    "Закройте его или укажите другое имя файла и повторите сохранение." % file_name,
                )
                return
            if not messagebox.askyesno(
                "Файл уже существует", "Файл «%s» уже есть на Рабочем столе.\nПерезаписать его?" % file_name
            ):
                return
        self._start_job("⏳ Конвертация и сохранение…", task_save_docx, self._after_save, md_text, dest)

    def _after_save(self, dest_path):
        try:
            open_with_default_app(dest_path)
            self._set_status("✅ Сохранено на Рабочий стол: %s" % os.path.basename(dest_path), "ok")
        except (OSError, AttributeError):
            self._set_status("✅ Сохранено: %s" % dest_path, "ok")

    def on_copy_text(self):
        md_text = self._get_text()
        if not md_text.strip():
            self._set_status("Поле пустое — вставьте Markdown.", "warn")
            return
        self._start_job("⏳ Обработка текста…", task_cure_markdown, self._after_copy_text, md_text)

    def _after_copy_text(self, cured):
        self._clipboard_set(cured)
        self._set_status(
            "✅ Вылеченный Markdown скопирован (%s символов)." % format(len(cured), ",").replace(",", " "), "ok"
        )

    # ---------------------------------------------------------------- закрытие
    def _on_close(self):
        self._closing = True
        try:
            self.root.update()   # даёт Tk отдать содержимое буфера обмена системе
        except tk.TclError:
            pass
        self.root.destroy()


def _enable_dpi_awareness():
    if not IS_WINDOWS:
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    _enable_dpi_awareness()
    root = tk.Tk()
    MathToWordApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
