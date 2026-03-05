"""
Detecta coords e salva resultado em arquivo UTF-8.
"""
import win32gui
import ctypes
import sys

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

output_lines = []

def log(msg):
    output_lines.append(msg)
    print(msg)


def encontrar_janela_sac():
    resultados = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            titulo = win32gui.GetWindowText(hwnd)
            if titulo and ("sistema de vendas" in titulo.lower()
                           or "benchimol" in titulo.lower()):
                rect = win32gui.GetWindowRect(hwnd)
                resultados.append((hwnd, titulo, rect))
    win32gui.EnumWindows(callback, None)
    return resultados


def listar_children(hwnd_top):
    children = []
    def cb(h, _):
        cls = win32gui.GetClassName(h)
        txt = win32gui.GetWindowText(h)
        rect = win32gui.GetWindowRect(h)
        cx = (rect[0] + rect[2]) // 2
        cy = (rect[1] + rect[3]) // 2
        children.append((h, cls, txt, rect, cx, cy))
    try:
        win32gui.EnumChildWindows(hwnd_top, cb, None)
    except Exception:
        pass
    children.sort(key=lambda e: (e[3][1], e[3][0]))
    return children


def main():
    user32 = ctypes.windll.user32
    w_primary = user32.GetSystemMetrics(0)
    h_primary = user32.GetSystemMetrics(1)

    log("Monitor primario: %dx%d" % (w_primary, h_primary))

    janelas = encontrar_janela_sac()
    if not janelas:
        log("ERRO: Janela SAC nao encontrada!")
        return

    hwnd_sac, titulo_sac, rect_sac = janelas[0]
    monitor = "DIREITO" if rect_sac[0] >= w_primary else "ESQUERDO"
    log("Janela: %s" % titulo_sac[:60])
    log("Monitor: %s" % monitor)
    log("Rect: left=%d top=%d right=%d bottom=%d" % rect_sac)

    children = listar_children(hwnd_sac)

    # Filtrar edits e botoes
    edits = [(h, cls, txt, rect, cx, cy) for h, cls, txt, rect, cx, cy in children
             if "edit" in cls.lower()]
    botoes = [(h, cls, txt, rect, cx, cy) for h, cls, txt, rect, cx, cy in children
              if "button" in cls.lower()]

    log("")
    log("=== EDITS (%d) ===" % len(edits))
    for i, (h, cls, txt, rect, cx, cy) in enumerate(edits):
        log("[%d] cls=%-18s txt='%-25s' center=(%d,%d)" % (i, cls, txt[:25], cx, cy))

    log("")
    log("=== BOTOES (%d) ===" % len(botoes))
    for h, cls, txt, rect, cx, cy in botoes:
        log("cls=%-18s txt='%-25s' center=(%d,%d)" % (cls, txt[:25], cx, cy))

    # Salvar resultado
    with open("coords_resultado.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
    log("")
    log("Resultado salvo em coords_resultado.txt")


if __name__ == "__main__":
    main()
