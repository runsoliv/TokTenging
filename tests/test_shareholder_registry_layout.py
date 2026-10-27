import unittest

from tok.shareholder_registry import (
    _capture_window_layout,
    _restore_window_layout,
)


class _FakeWindow:
    def __init__(self, state="normal"):
        self._state = state
        self._width = 680
        self._height = 700
        self._x = 120
        self._y = 80
        self._minsize = (620, 560)
        self.geometry_calls = []

    def update_idletasks(self):
        return None

    def state(self, value=None):
        if value is not None:
            self._state = value
        return self._state

    def winfo_width(self):
        return self._width

    def winfo_height(self):
        return self._height

    def winfo_x(self):
        return self._x

    def winfo_y(self):
        return self._y

    def minsize(self, width=None, height=None):
        if width is not None and height is not None:
            self._minsize = (width, height)
        return self._minsize

    def geometry(self, value):
        self.geometry_calls.append(value)


class RegistryWindowLayoutTests(unittest.TestCase):
    def test_normal_window_returns_to_exact_previous_layout(self):
        window = _FakeWindow()
        layout = _capture_window_layout(window)
        window._width = 1120
        window._height = 790
        window._x = 24
        window._y = 16
        window._minsize = (880, 600)

        _restore_window_layout(window, layout)

        self.assertEqual(window.geometry_calls, ["680x700+120+80"])
        self.assertEqual(window.minsize(), (620, 560))
        self.assertEqual(window.state(), "normal")

    def test_maximized_window_returns_to_maximized_state(self):
        window = _FakeWindow(state="zoomed")
        layout = _capture_window_layout(window)
        window._state = "normal"
        window._minsize = (880, 600)

        _restore_window_layout(window, layout)

        self.assertEqual(window.state(), "zoomed")
        self.assertEqual(window.minsize(), (620, 560))
        self.assertEqual(window.geometry_calls, [])


if __name__ == "__main__":
    unittest.main()
