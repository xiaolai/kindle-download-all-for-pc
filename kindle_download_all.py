import argparse
import time
from typing import List, Optional, Tuple

from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
from pywinauto.timings import TimeoutError as WaitTimeoutError
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_defines import IUIA

KINDLE_EXECUTABLE = "Kindle.exe"
MAIN_WINDOW_TITLE = "Kindle"
PREFERRED_LIST_IDS: List[str] = ["compact", "row"]
DOWNLOAD_MENU_AUTO_ID = "context-menu-option-download-book"
MENU_TITLE_KEYWORDS = ["\u4e0b\u8f7d", "Download"]

DELAY_BETWEEN_ACTIONS = 0.5
MENU_APPEAR_TIMEOUT = 2.0
FOCUS_SETTLE_DELAY = 0.3
CONTEXT_MENU_DELAY = 0.3
KEYBOARD_SELECTION_DELAY = 0.5
ARROW_MOVE_DELAY = 0.2
DEFAULT_MAX_ITERATIONS = 10000


def as_wrapper(control):
    """Return a UIA wrapper for the given control specification."""
    try:
        return control.wrapper_object()
    except AttributeError:
        return control


def visible_items(library) -> List:
    """Return visible list items from the Kindle library control."""
    wrapper = as_wrapper(library)
    try:
        descendants = wrapper.descendants(control_type="ListItem")
    except Exception:
        return []
    return [item for item in descendants if getattr(item.element_info, "visible", False)]


def focus_item(item) -> None:
    """Bring the given list item into focus using selection or mouse."""
    wrapper = as_wrapper(item)
    try:
        wrapper.scroll_into_view()
    except Exception:
        pass
    try:
        wrapper.set_focus()
    except Exception:
        try:
            wrapper.select()
        except Exception:
            wrapper.click_input()


def item_identity(item) -> Tuple:
    """Produce a stable identity tuple for a list item."""
    wrapper = as_wrapper(item)
    info = wrapper.element_info
    runtime_id = getattr(info, "runtime_id", None)
    if runtime_id:
        return ("runtime", tuple(runtime_id))
    return (
        "fallback",
        info.name,
        getattr(info, "automation_id", None),
        getattr(info, "control_type", None),
    )


def safe_text(text: str) -> str:
    """Return text encoded for the current console without errors."""
    try:
        return text.encode("gbk", errors="replace").decode("gbk")
    except Exception:
        return text


def get_main_window(app: Application):
    """Return the Kindle main window, using title or best match as fallback."""
    try:
        window = app.window(title=MAIN_WINDOW_TITLE)
        window.wait("visible", timeout=5)
        return window
    except Exception:
        return app.window(best_match="Kindle")


def locate_library(main_window):
    """Locate the Kindle library list control with fallbacks."""
    for auto_id in PREFERRED_LIST_IDS:
        try:
            spec = main_window.child_window(auto_id=auto_id, control_type="List")
            wrapper = as_wrapper(spec)
            if visible_items(wrapper):
                return wrapper
        except ElementNotFoundError:
            continue
        except Exception:
            continue

    candidates = main_window.descendants(control_type="List")
    best_match = None
    best_count = 0
    for ctrl in candidates:
        try:
            wrapper = as_wrapper(ctrl)
            items = visible_items(wrapper)
            if len(items) > best_count:
                best_match = wrapper
                best_count = len(items)
        except Exception:
            continue

    if best_match and best_count:
        return best_match

    raise ElementNotFoundError("Could not locate a visible List control for the library.")


def find_download_menu_item(timeout: float):
    """Locate the download menu item using automation id or known titles."""
    desktop = Desktop(backend="uia")
    search_specs = [
        dict(auto_id=DOWNLOAD_MENU_AUTO_ID, control_type="MenuItem", top_level_only=False),
    ]
    for keyword in MENU_TITLE_KEYWORDS:
        search_specs.append(dict(control_type="MenuItem", title=keyword, top_level_only=False))
        search_specs.append(dict(control_type="MenuItem", title_re=f".*{keyword}.*", top_level_only=False))

    deadline = time.time() + timeout
    while time.time() < deadline:
        for spec_kwargs in search_specs:
            try:
                spec = desktop.window(**spec_kwargs)
            except Exception:
                continue

            try:
                spec.wait("exists enabled visible ready", timeout=0.1)
                return as_wrapper(spec)
            except WaitTimeoutError:
                continue
            except Exception:
                continue
        time.sleep(0.05)

    return None


def open_context_menu() -> bool:
    """Open the context menu using keyboard shortcuts."""
    send_keys("+{F10}")
    time.sleep(CONTEXT_MENU_DELAY)
    return find_download_menu_item(MENU_APPEAR_TIMEOUT) is not None


def trigger_download_command() -> bool:
    """Activate the download command via automation id."""
    download_item = find_download_menu_item(timeout=0.5)
    if download_item:
        try:
            download_item.invoke()
        except Exception:
            try:
                download_item.click_input()
            except Exception:
                send_keys("{ESC}")
                return False
        print("  - Download command triggered.")
        return True
    return False


def get_focused_list_item() -> Optional[UIAWrapper]:
    """Return the currently focused list item, if any."""
    try:
        focused = IUIA().get_focused_element()
        if focused:
            wrapper = UIAWrapper(focused)
            if wrapper.element_info.control_type == "ListItem":
                return wrapper
    except Exception:
        return None
    return None


def get_total_items(library) -> Optional[int]:
    try:
        return library.item_count()
    except Exception:
        return None


def find_start_index(library, target_wrapper, total_items: Optional[int]) -> int:
    if target_wrapper is None:
        return 0

    target_id = item_identity(target_wrapper)
    if total_items is None:
        idx = 0
        while idx < DEFAULT_MAX_ITERATIONS:
            try:
                candidate = library.get_item(idx)
            except Exception:
                break
            if item_identity(candidate) == target_id:
                return idx
            idx += 1
        return 0

    for idx in range(total_items):
        try:
            candidate = library.get_item(idx)
        except Exception:
            break
        if item_identity(candidate) == target_id:
            return idx
    return 0


def iterate_items(library, start_index: int, total_items: Optional[int], max_iterations: int) -> None:
    processed_count = 0
    index = start_index
    last_identity: Optional[Tuple] = None

    while processed_count < max_iterations:
        try:
            current_item = library.get_item(index)
        except Exception:
            print("Reached end of list (unable to retrieve more items).")
            break

        current_id = item_identity(current_item)
        if last_identity is not None and current_id == last_identity:
            print(f"Reached end of list (duplicate item at {index + 1}).")
            break

        title = safe_text(current_item.window_text())
        processed_count += 1
        current_number = index + 1
        print(f"[{current_number}] Processing '{title}'")

        focus_item(current_item)
        time.sleep(FOCUS_SETTLE_DELAY)

        if not open_context_menu():
            print("  ! Context menu did not appear; skipping this item.")
        else:
            if not trigger_download_command():
                print("  ! Download option not found; skipping this item.")
            time.sleep(DELAY_BETWEEN_ACTIONS)

        last_identity = current_id
        index += 1
        send_keys("{DOWN}")
        time.sleep(ARROW_MOVE_DELAY)

    if processed_count >= max_iterations:
        print(f"Reached iteration limit ({max_iterations}); stopping at list index {index}.")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Trigger Kindle downloads from the current selection to the end of the list."
    )
    parser.add_argument(
        "--start-from",
        type=int,
        metavar="INDEX",
        help="1-based index of the book to start from (defaults to currently focused item).",
    )
    parser.add_argument(
        "--max-iterations",
        type=int,
        default=DEFAULT_MAX_ITERATIONS,
        help="Maximum number of books to process in one run (default: %(default)s).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        app = Application(backend="uia").connect(path=KINDLE_EXECUTABLE)
    except ElementNotFoundError:
        raise SystemExit("Kindle for PC is not running or cannot be found.")

    main_window = get_main_window(app)
    try:
        library = locate_library(main_window)
    except ElementNotFoundError as exc:
        raise SystemExit(str(exc))

    total_items = get_total_items(library)

    if args.start_from is not None:
        start_index = max(args.start_from - 1, 0)
        if total_items is not None and start_index >= total_items:
            print(
                f"Start index {start_index + 1} exceeds total reported items ({total_items}). Adjusting to last item."
            )
            start_index = max(total_items - 1, 0)
    else:
        focused_wrapper = get_focused_list_item()
        if focused_wrapper is None:
            library.set_focus()
            time.sleep(FOCUS_SETTLE_DELAY)
            focused_wrapper = get_focused_list_item()
        start_index = find_start_index(library, focused_wrapper, total_items)

    if total_items is not None:
        print(f"Library reports {total_items} items; starting at index {start_index} (1-based {start_index + 1}).")
    else:
        print(f"Starting at index {start_index} (1-based {start_index + 1}); total count not available.")

    iterate_items(library, start_index, total_items, args.max_iterations)


if __name__ == "__main__":
    main()
