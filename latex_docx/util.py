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
    """Who triggered the operation: agent | user (from flag, env, or default)."""
    if cli_by:
        return cli_by.strip().lower()
    env = (os.environ.get("LATEX_DOCX_BY") or "").strip().lower()
    if env in ("agent", "user"):
        return env
    # Cursor / agent sessions often set one of these
    if any(os.environ.get(k) for k in ("CURSOR_AGENT", "CURSOR_TRACE_ID", "COMPOSER_SESSION_ID")):
        return "agent"
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



