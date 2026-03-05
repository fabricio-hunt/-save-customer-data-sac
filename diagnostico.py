"""
DIAGNÓSTICO - Bemol SAC (sem psutil)
Execute: python diagnostico.py
"""
import os
import win32gui
import win32process
import win32con

# ── 1. Tipo de sessão ────────────────────────────────────────────────────────
print("=== SESSÃO ATUAL ===")
sessao = os.environ.get('SESSIONNAME', 'N/A')
print(f"  SESSIONNAME : {sessao}")
print(f"  USERNAME    : {os.environ.get('USERNAME','?')}")
print(f"  COMPUTERNAME: {os.environ.get('COMPUTERNAME','?')}")
if 'rdp' in sessao.lower() or 'tcp' in sessao.lower():
    print("  >> SESSAO REMOTA (RDP)")
else:
    print("  >> Sessao LOCAL (Console)")

# ── 2. Janelas com campos Edit ───────────────────────────────────────────────
print("\n=== JANELAS COM CAMPOS EDIT ===")

def coletar_edits(hwnd_top):
    edits = []
    def cb(h, _):
        if win32gui.GetClassName(h) == "Edit":
            txt  = win32gui.GetWindowText(h)
            rect = win32gui.GetWindowRect(h)
            edits.append((h, txt, rect))
    try:
        win32gui.EnumChildWindows(hwnd_top, cb, None)
    except Exception:
        pass
    return edits

def enum_top(hwnd, _):
    if not win32gui.IsWindowVisible(hwnd):
        return
    titulo = win32gui.GetWindowText(hwnd)
    if not titulo:
        return
    edits = coletar_edits(hwnd)
    if edits:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        print(f"\nJanela : '{titulo}'")
        print(f"  hwnd={hwnd}  pid={pid}")
        edits.sort(key=lambda e: (e[2][1], e[2][0]))  # ordena por y,x
        for i, (h, txt, rect) in enumerate(edits):
            print(f"  [{i}] hwnd={h:<10} texto='{txt:<20}' pos=({rect[0]},{rect[1]})")

win32gui.EnumWindows(enum_top, None)

# ── 3. Todas as janelas visíveis ─────────────────────────────────────────────
print("\n\n=== TODAS AS JANELAS VISÍVEIS (top-level) ===")
print(f"{'HWND':<12} {'PID':<8} Título")
print("-"*60)

def enum_all(hwnd, _):
    if not win32gui.IsWindowVisible(hwnd):
        return
    titulo = win32gui.GetWindowText(hwnd)
    if not titulo:
        return
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        pid = -1
    print(f"{hwnd:<12} {pid:<8} {titulo[:55]}")

win32gui.EnumWindows(enum_all, None)
print("\nFim. Cole tudo acima.")