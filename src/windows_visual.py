"""Identidade e nitidez do aplicativo no Windows, antes de criar a janela."""
import os
import sys


def preparar_windows():
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MarineSeguros.MarinePGR")
    except (AttributeError, OSError):
        pass
    try:
        # Tk 8.6 funciona melhor com escala de sistema do que por monitor.
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def aplicar_icone(janela, caminho):
    janela.iconbitmap(default=os.path.abspath(caminho))
    if sys.platform != "win32":
        return
    import ctypes
    from ctypes import wintypes
    janela.update_idletasks()
    user = ctypes.windll.user32
    user.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user.GetAncestor.restype = wintypes.HWND
    user.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
                               ctypes.c_int, ctypes.c_int, wintypes.UINT]
    user.LoadImageW.restype = wintypes.HANDLE
    user.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user.SendMessageW.restype = ctypes.c_ssize_t
    hwnd = user.GetAncestor(janela.winfo_id(), 2) or janela.winfo_id()
    try:
        user.GetDpiForWindow.argtypes = [wintypes.HWND]
        user.GetDpiForWindow.restype = wintypes.UINT
        dpi = user.GetDpiForWindow(hwnd) or 96
    except AttributeError:
        dpi = round(janela.winfo_fpixels("1i"))
    handles = []
    for tipo, tamanho in ((0, round(16 * dpi / 96)), (1, round(32 * dpi / 96))):
        handle = user.LoadImageW(None, os.path.abspath(caminho), 1, tamanho, tamanho, 0x10)
        if handle:
            user.SendMessageW(hwnd, 0x80, tipo, handle)
            handles.append(handle)
    janela._marine_icon_handles = handles
    def liberar(evento):
        if evento.widget is janela:
            user.DestroyIcon.argtypes = [wintypes.HICON]
            for handle in handles:
                user.DestroyIcon(handle)
    janela.bind("<Destroy>", liberar, add="+")
