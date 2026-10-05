"""Profile Manager layout: controls stay reachable at the minimum window size."""
from __future__ import annotations

import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture(scope="module")
def manager():
    from gui import profile_manager as pm

    tk.Tk.mainloop = lambda self: None
    try:
        instance = pm.ProfileManager()
    except tk.TclError:
        pytest.skip("No display available for the Profile Manager window")
    instance.root.geometry("1180x780")  # minimum supported size
    instance.root.update_idletasks()
    yield pm, instance
    instance.root.destroy()


def _visible(widget, root) -> bool:
    try:
        x, y, w, h = widget.winfo_rootx(), widget.winfo_rooty(), widget.winfo_width(), widget.winfo_height()
    except Exception:
        return False
    return w > 0 and h > 0 and (y + h) <= (root.winfo_rooty() + root.winfo_height()) + 2 and (x + w) <= (root.winfo_rootx() + root.winfo_width()) + 2


def _find_button(widget, text):
    for child in widget.winfo_children():
        try:
            if child.cget("text") == text:
                return child
        except Exception:
            pass
        found = _find_button(child, text)
        if found is not None:
            return found
    return None


def test_bottom_actions_and_status_are_never_clipped(manager):
    pm, app = manager
    assert _visible(app.install_button.master, app.root)  # footer with Save/Validate/Install
    assert _visible(app.back_step_button.master, app.root)  # step Back/Next bar
    assert _visible(app.root.pack_slaves()[-1], app.root)  # status bar
    delete = _find_button(app.catalog.master, "Delete…")
    assert delete is not None and _visible(delete, app.root)


def test_every_step_page_scrolls_so_content_is_reachable(manager):
    pm, app = manager
    for key in pm.STEPS:
        app._show_step(key)
        app.root.update_idletasks()
        host = app.pages[key]._scroll_host
        assert host.canvas.bbox("all") is not None
        assert _visible(app.next_step_button, app.root)
        assert _visible(app.back_step_button, app.root)
    app._show_step("overview")
