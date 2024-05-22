import pythoncom # should be imported before pywinauto
from pywinauto import findwindows, Desktop, Application

# find the window
# backend = GUI framework used to develop the application
windows = findwindows.find_elements() # find all windows
for window in windows:
    window_name = window.name
    if '원격 데스크톱 연결' in window_name:
        app = Application(backend='win32').connect(title=window_name)

dlg = app.window(title_re='원격 데스크톱 연결')



# If you want to run automatically even with disconnection from remote host.
# https://stackoverflow.com/questions/50299472/pywin32-pywinauto-not-working-properly-in-remote-desktop-when-it-is-minimized
# https://pywinauto.readthedocs.io/en/latest/remote_execution.html

