from __future__ import annotations

from pathlib import Path
from typing import Callable

from PyQt6.QtCore import QTimer
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import QPoint
from PyQt6.QtWidgets import QApplication, QMainWindow, QToolBar, QToolButton, QWidget


ACTION_OBJECT_NAME = "llIntegrationExperimentalToolbarAction"
ACTION_MARKER = "_ll_integration_toolbar_action"


def install_toolbar_button(
    main_window,
    icon_path: Path,
    callback: Callable[[], None],
    log: Callable[[str], None] | None = None,
) -> None:
    """Try to add a direct LL Integration action to MO2's main toolbar.

    This is intentionally best-effort. MO2 does not expose an official Python API
    for pinning plugin tools to the main toolbar, so this uses Qt widget discovery.
    """

    def write_log(message: str) -> None:
        if log:
            log(message)

    def current_main_window():
        window = main_window() if callable(main_window) else main_window
        if window is not None:
            return window
        return _find_application_main_window(write_log)

    def attempt() -> bool:
        window = current_main_window()
        write_log("toolbar experiment: attempt")
        if window is None:
            write_log("toolbar experiment: no main window")
            return False

        if hasattr(window, ACTION_MARKER):
            write_log("toolbar experiment: action already installed")
            return True

        toolbar = _find_toolbar(window)
        if toolbar is None:
            write_log("toolbar experiment: no QToolBar found")
            _dump_widget_hints(window, write_log)
            return False

        action = QAction(QIcon(str(icon_path)), "LL Integration", toolbar)
        action.setObjectName(ACTION_OBJECT_NAME)
        action.setToolTip("Open LL Integration")
        action.triggered.connect(lambda _checked=False: callback())
        before_action = _find_insert_before_action(toolbar, write_log)
        if before_action is not None:
            toolbar.insertAction(before_action, action)
        else:
            toolbar.addAction(action)
        setattr(window, ACTION_MARKER, action)
        write_log(
            "toolbar experiment: added action to "
            f"{toolbar.objectName() or '<unnamed toolbar>'} "
            f"actions={len(toolbar.actions())} visible={toolbar.isVisible()} "
            f"before={before_action.text() if before_action else '<end>'}"
        )
        return True

    if attempt():
        return

    # MO2 can finish wiring its main window shortly after plugin init.
    for delay_ms in (250, 1000, 2500):
        QTimer.singleShot(delay_ms, attempt)


def _find_toolbar(main_window) -> QToolBar | None:
    if isinstance(main_window, QMainWindow):
        toolbars = main_window.findChildren(QToolBar)
    else:
        toolbars = main_window.findChildren(QToolBar) if hasattr(main_window, "findChildren") else []

    visible = [toolbar for toolbar in toolbars if toolbar.isVisible()]
    candidates = visible or toolbars
    if not candidates:
        _find_tool_buttons(main_window)
        return None

    # Prefer the toolbar that already contains actions, which is usually MO2's main toolbar.
    candidates.sort(key=lambda toolbar: len(toolbar.actions()), reverse=True)
    for toolbar in candidates:
        if any(action.objectName() == ACTION_OBJECT_NAME for action in toolbar.actions()):
            return toolbar
    return candidates[0]


def _find_application_main_window(write_log: Callable[[str], None]):
    app = QApplication.instance()
    if app is None:
        write_log("toolbar experiment: no QApplication instance")
        return None

    widgets = app.topLevelWidgets()
    write_log(f"toolbar experiment: top-level widgets={len(widgets)}")
    main_windows = [widget for widget in widgets if isinstance(widget, QMainWindow)]
    if main_windows:
        main_windows.sort(key=lambda widget: int(widget.isVisible()), reverse=True)
        chosen = main_windows[0]
        write_log(
            "toolbar experiment: found QMainWindow "
            f"title={chosen.windowTitle() or '<none>'} "
            f"name={chosen.objectName() or '<none>'} "
            f"visible={chosen.isVisible()}"
        )
        return chosen

    for widget in widgets[:20]:
        write_log(
            "toolbar experiment: top-level "
            f"{type(widget).__name__} "
            f"title={widget.windowTitle() or '<none>'} "
            f"name={widget.objectName() or '<none>'} "
            f"visible={widget.isVisible()}"
        )

    return None


def _find_insert_before_action(toolbar: QToolBar, write_log: Callable[[str], None]) -> QAction | None:
    actions = [action for action in toolbar.actions() if action.isVisible()]
    if not actions:
        return None

    toolbar_width = max(toolbar.width(), 1)
    button_positions = []
    for action in actions:
        widget = toolbar.widgetForAction(action)
        if widget is None or not widget.isVisible():
            continue

        center = widget.mapTo(toolbar, QPoint(widget.width() // 2, widget.height() // 2))
        button_positions.append((center.x(), action, widget))

    if not button_positions:
        return None

    button_positions.sort(key=lambda item: item[0])
    write_log(
        "toolbar experiment: action positions "
        + " | ".join(
            f"{action.text() or action.objectName() or type(widget).__name__}@{x}"
            for x, action, widget in button_positions
        )
    )

    # Insert before the first cluster that starts well past the center. In MO2 this
    # puts LL Integration at the beginning of the right-side icon group instead of
    # after every help/status action.
    for x, action, _widget in button_positions:
        if x > toolbar_width * 0.55:
            return action

    return None


def _find_tool_buttons(main_window) -> list[QToolButton]:
    if not hasattr(main_window, "findChildren"):
        return []
    try:
        return main_window.findChildren(QToolButton)
    except Exception:
        return []


def _dump_widget_hints(main_window, write_log: Callable[[str], None]) -> None:
    if not hasattr(main_window, "findChildren"):
        return

    try:
        widgets = main_window.findChildren(QWidget)
    except Exception as exc:
        write_log(f"toolbar experiment: widget dump failed: {exc}")
        return

    interesting = []
    tool_buttons = set(_find_tool_buttons(main_window))
    for widget in widgets:
        class_name = type(widget).__name__
        object_name = widget.objectName() if hasattr(widget, "objectName") else ""
        is_tool_button = widget in tool_buttons
        if (
            not is_tool_button
            and "tool" not in class_name.lower()
            and "tool" not in object_name.lower()
            and "bar" not in object_name.lower()
        ):
            continue

        try:
            geometry = widget.geometry()
            geo = f"{geometry.x()},{geometry.y()} {geometry.width()}x{geometry.height()}"
        except Exception:
            geo = "unknown geometry"

        interesting.append(
            f"{class_name} name={object_name or '<none>'} "
            f"visible={widget.isVisible()} geo={geo}"
        )

    write_log(f"toolbar experiment: widget hints count={len(interesting)}")
    for line in interesting[:80]:
        write_log(f"toolbar experiment: widget {line}")
