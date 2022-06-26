import win32gui,win32api,win32con
import time
import random
 
class TestTaskbarIcon:
    def __init__(self):
        # 注册一个窗口类
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32gui.GetModuleHandle(None)
        name_str = str(random.uniform(0,1000000000000))#随机字符串，防止报错pywintypes.error: (1410, 'RegisterClass', '类已存在。')
        wc.lpszClassName = name_str
        wc.lpfnWndProc = {win32con.WM_DESTROY: self.OnDestroy, }
        classAtom = win32gui.RegisterClass(wc)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom, "Taskbar Demo", style,
                                          0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                                          0, 0, hinst, None)
        hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        self.hicon=hicon
        nid = (self.hwnd, 0, win32gui.NIF_ICON, win32con.WM_USER + 20, hicon, "Demo")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
 
    def showMsg(self, title, msg):
        # 原作者使用Shell_NotifyIconA方法代替包装后的Shell_NotifyIcon方法
        # 据称是不能win32gui structure, 我稀里糊涂搞出来了.
        # 具体对比原代码.
        nid = (self.hwnd,  # 句柄
               0,  # 托盘图标ID
               win32gui.NIF_INFO,  # 标识
               0,  # 回调消息ID
               0,  # 托盘图标句柄
               "TestMessage",  # 图标字符串
               msg,  # 气球提示字符串
               0,  # 提示的显示时间←这个提示时间改了也没用的样子
               title,  # 提示标题
               win32gui.NIIF_INFO  # 提示用到的图标
               )
        self.nid = nid
        win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

def show_msg(title,msg,seconds):
	t = TestTaskbarIcon()
	t.showMsg(title, msg)
	time.sleep(seconds)
	win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE,t.nid)#删除托盘图标。如果用其他的方法win32api.MessageBox就弹不出来

	##以下都是失败的退出方法
	# win32gui.PostQuitMessage(0) 
	# win32gui.DestroyWindow(t.hwnd)
	# win32gui.PostMessage(t.hwnd, win32con.WM_NULL, 0, 0) 

