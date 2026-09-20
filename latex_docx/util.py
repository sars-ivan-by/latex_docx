"""Paths, logging, TeX brace helpers."""
from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Union

from .constants import DEFAULT_OP_LOG

def ensure_out_prefix(path: Union[str, Path]) -> Path:
    """Гарантировать приставку out в имени файла (для .gitignore: out*)."""
    p = Path(path)
    if p.name.startswith("out"):
        return p
    return p.with_name(f"out_{p.name}")



def resolve_by(cli_by: Optional[str] = None) -> str:
    """Who triggered the operation: agent | user (from flag, env, or heuristics)."""
    if cli_by:
        return cli_by.strip().lower()
    env = (os.environ.get("LATEX_DOCX_BY") or "").strip().lower()
    if env in ("agent", "user"):
        return env

    # Explicit Cursor / agent session markers
    agent_env_keys = (
        "CURSOR_AGENT",
        "CURSOR_AGENT_ID",
        "CURSOR_TRACE_ID",
        "CURSOR_SESSION_ID",
        "COMPOSER_SESSION_ID",
        "COMPOSER_GENERATION_ID",
        "CLAUDE_AGENT",
        "AIDER_AGENT",
        "AI_AGENT",
    )
    if any(os.environ.get(k) for k in agent_env_keys):
        return "agent"

    # Cursor CLI / cloud agent often sets these
    if (os.environ.get("CURSOR") or "").strip().lower() in ("1", "true", "yes"):
        return "agent"
    if os.environ.get("CURSOR_API_KEY") and os.environ.get("CI"):
        return "agent"

    # Parent process / argv heuristics (best-effort, never raise)
    try:
        import sys

        joined = " ".join(sys.argv).lower()
        if any(tok in joined for tok in ("cursor-agent", "composer", "/agent", "claude-code")):
            return "agent"
    except Exception:
        pass

    try:
        import psutil  # optional

        proc = psutil.Process(os.getpid())
        for p in proc.parents():
            name = (p.name() or "").lower()
            if any(x in name for x in ("cursor", "composer", "claude", "codex", "agent")):
                return "agent"
    except Exception:
        pass

    return "user"



def append_op_log(
    action: str,
    details: str,
    *,
    by: Optional[str] = None,
    log_path: Optional[Union[str, Path]] = None,
) -> Path:
    """Append one line to out_latex_docx.log (who / what / when)."""
    path = ensure_out_prefix(log_path) if log_path else DEFAULT_OP_LOG
    who = resolve_by(by)
    ts = datetime.now().astimezone().strftime("%Y-%m-%dT%H:%M:%S%z")
    line = f"{ts} | {action:<12} | by={who:<5} | {os.environ.get('USER') or os.environ.get('LOGNAME') or '-'} | cwd={Path.cwd()} | {details}\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(line)
    return path


GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ",
    "phi": "φ", "varphi": "ϕ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Pi": "Π",
    "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}
MATH_OPS = {
    "cdot": "·", "times": "×", "pm": "±", "leq": "≤", "geq": "≥", "neq": "≠",
    "approx": "≈", "propto": "∝", "infty": "∞", "partial": "∂", "sum": "∑",
    "prod": "∏", "int": "∫", "rightarrow": "→", "leftarrow": "←", "to": "→",
    "ldots": "…", "cdots": "⋯", "dots": "…", "circ": "∘", "mod": "mod",
    "bmod": "mod", "quad": " ", "qquad": "  ", "ell": "ℓ", "Re": "Re", "Im": "Im",
    "subseteq": "⊆", "in": "∈", "forall": "∀", "exists": "∃", "nabla": "∇",
}



def strip_comments(tex: str) -> str:
    out = []
    for line in tex.splitlines():
        i, buf = 0, []
        while i < len(line):
            if line[i] == "%" and (i == 0 or line[i - 1] != "\\"):
                break
            buf.append(line[i])
            i += 1
        out.append("".join(buf))
    return "\n".join(out)



def find_balanced(s: str, start: int) -> Tuple[str, int]:
    if start >= len(s) or s[start] != "{":
        raise ValueError(f"Expected '{{' at {start}")
    depth, i = 0, start
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return s[start + 1 : i], i + 1
        i += 1
    raise ValueError("Unbalanced braces")



def extract_command_arg(tex: str, cmd: str) -> Optional[str]:
    m = re.search(rf"\\{cmd}\s*\{{", tex)
    if not m:
        return None
    arg, _ = find_balanced(tex, m.end() - 1)
    return arg



def extract_environment(tex: str, name: str, start: int = 0):
    begin = re.search(rf"\\begin\{{{name}\}}", tex[start:])
    if not begin:
        return None
    abs_begin = start + begin.start()
    body_start = start + begin.end()
    depth, i = 1, body_start
    pat_b = re.compile(rf"\\begin\{{{name}\}}")
    pat_e = re.compile(rf"\\end\{{{name}\}}")
    while i < len(tex) and depth > 0:
        mb, me = pat_b.search(tex, i), pat_e.search(tex, i)
        if me is None:
            return None
        if mb and mb.start() < me.start():
            depth += 1
            i = mb.end()
        else:
            depth -= 1
            if depth == 0:
                return tex[body_start : me.start()], abs_begin, me.end()
            i = me.end()
    return None



def _safe_balanced(s: str, start: int):
    try:
        return find_balanced(s, start)
    except ValueError:
        return None



def expand_inputs(tex: str, base_dir: Path, *, _stack: Optional[set] = None) -> str:
    """Expand \\input{...} / \\include{...} recursively (cycle-safe)."""
    if _stack is None:
        _stack = set()

    def replacer(m: re.Match) -> str:
        rel = m.group(1).strip()
        if not rel:
            return ""
        cand = Path(rel)
        if not cand.suffix:
            cand = cand.with_suffix(".tex")
        path = (base_dir / cand).resolve()
        key = str(path)
        if key in _stack:
            return f"% cyclic \\input{{{rel}}} skipped\n"
        if not path.exists():
            # Try as-is without forcing .tex
            alt = (base_dir / rel).resolve()
            if alt.exists():
                path = alt
                key = str(path)
            else:
                return m.group(0)  # leave unresolved
        _stack.add(key)
        try:
            nested = strip_comments(path.read_text(encoding="utf-8", errors="replace"))
            return expand_inputs(nested, path.parent, _stack=_stack)
        finally:
            _stack.discard(key)

    # \input{file} and \include{file} (not \includegraphics)
    tex = re.sub(r"\\include\s*\{([^}]+)\}", replacer, tex)
    tex = re.sub(r"\\input\s*\{([^}]+)\}", replacer, tex)
    return tex



def rasterize_pdf(pdf_path: Path, *, dpi: int = 150) -> Optional[Path]:
    """Render first page of a PDF to PNG. Uses pdftoppm / magick when available."""
    import shutil
    import subprocess
    import tempfile

    pdf_path = Path(pdf_path)
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        return None

    cache = pdf_path.with_name(pdf_path.stem + f"_p1_{dpi}.png")
    if cache.exists() and cache.stat().st_mtime >= pdf_path.stat().st_mtime:
        return cache

    tmp_dir = Path(tempfile.mkdtemp(prefix="latex_docx_pdf_"))
    try:
        out_base = tmp_dir / "page"
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            subprocess.run(
                [pdftoppm, "-png", "-r", str(dpi), "-f", "1", "-l", "1", str(pdf_path), str(out_base)],
                check=True,
                capture_output=True,
            )
            produced = out_base.with_name(out_base.name + "-1.png")
            if produced.exists():
                cache.write_bytes(produced.read_bytes())
                return cache

        magick = shutil.which("magick") or shutil.which("convert")
        if magick:
            tmp_png = tmp_dir / "page.png"
            subprocess.run(
                [magick, "-density", str(dpi), f"{pdf_path}[0]", str(tmp_png)],
                check=True,
                capture_output=True,
            )
            if tmp_png.exists():
                cache.write_bytes(tmp_png.read_bytes())
                return cache
    except Exception:
        return None
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return None



