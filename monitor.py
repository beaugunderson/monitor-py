from win32api import GetCurrentProcess, OpenProcess, CloseHandle
from win32con import PROCESS_QUERY_INFORMATION, PROCESS_VM_READ, WM_GETTEXT
from win32gui import FindWindowEx, PyMakeBuffer, SendMessage, \
        GetForegroundWindow, GetWindowText
from win32process import GetWindowThreadProcessId, GetModuleFileNameEx
from win32security import OpenProcessToken, LookupPrivilegeValue, \
        AdjustTokenPrivileges, TOKEN_ADJUST_PRIVILEGES, \
        TOKEN_QUERY, SE_DEBUG_NAME, SE_PRIVILEGE_ENABLED

import codecs
import json
import sys
import time

def setup_cli_encoding():
    # tolerant unicode output ... #
    _stdout = sys.stdout

    if sys.platform == 'win32' and not \
        'pywin.framework.startup' in sys.modules:
        _stdoutenc = getattr(_stdout, 'encoding', sys.getdefaultencoding())

        class StdOut:
            def write(self,s):
                _stdout.write(s.encode(_stdoutenc, 'backslashreplace'))

        sys.stdout = StdOut()
    elif sys.platform.startswith('linux'):
        import locale

        _stdoutenc = locale.getdefaultlocale()[1]

        class StdOut:
            def write(self,s):
                _stdout.write(s.encode(_stdoutenc, 'backslashreplace'))

        sys.stdout = StdOut()

def escalate_privileges():
    # Request privileges to enable "debug process", so we can
    # later use PROCESS_VM_READ, required to GetModuleFileNameEx()
    privilege_flags = TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY

    process_token = OpenProcessToken(GetCurrentProcess(), privilege_flags)

    # enable "debug process"
    privilege_id = LookupPrivilegeValue(None, SE_DEBUG_NAME)

    AdjustTokenPrivileges(process_token, 0,
            [(privilege_id, SE_PRIVILEGE_ENABLED)])

    return process_token

def exe_from_window(w):
    process_id = GetWindowThreadProcessId(w)
    process_handle = OpenProcess(PROCESS_QUERY_INFORMATION | \
            PROCESS_VM_READ, False, process_id[1])

    exe = GetModuleFileNameEx(process_handle, 0)

    CloseHandle(process_handle)

    return exe

def url_from_chrome(w):
    view = FindWindowEx(w, 0, "Chrome_AutocompleteEditView", None)

    MAX_CHARACTERS = 1024

    buffer = PyMakeBuffer(MAX_CHARACTERS)
    length = SendMessage(view, WM_GETTEXT, MAX_CHARACTERS, buffer)

    result = buffer[:length]

    return result

def main():
    setup_cli_encoding()

    process_token = escalate_privileges()

    updates = []

    try:
        while True:
            handle = GetForegroundWindow()
            title = GetWindowText(handle)

            exe = exe_from_window(handle)

            update = {
                'utc_time': time.gmtime(),
                'executable': exe,
                'title': title
            }

            if exe.endswith("chrome.exe"):
                update['url'] = url_from_chrome(handle)

            updates.append(update)

            time.sleep(5)
    except KeyboardInterrupt:
        pass

    CloseHandle(process_token)

    print json.dumps(updates, indent=3)

if __name__ == "__main__":
    main()
