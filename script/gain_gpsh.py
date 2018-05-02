# coding=utf-8
import win32gui
import win32con
import cv2
import aircv as ac
import pytesseract
from PIL import ImageGrab

class ScriptForRadarView():
    def __init__(self):
        self.handle_tool = Handle()
        self.win = Windows()
        self.cV = OpenCv()
        self.radar_view_handle = None

    def gain_GPS_handle(self):
        MDIClient_handle = self.get_handle('MDIClient')
        # todo 这里会有角标越界的bug
        view_path_name = str(self.radar_view_name).split(" - ")[1]
        hwndChildList = []
        handle = None
        win32gui.EnumChildWindows(MDIClient_handle, lambda hwnd, param: param.append(hwnd), hwndChildList)
        for hwnd in hwndChildList:
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                window_name = win32gui.GetWindowText(hwnd)
                if window_name and str(view_path_name).find(window_name) >= 0:
                    handle = hwnd
                    break
        AfxMDIFrame42s_1 = win32gui.FindWindowEx(handle, None, 'AfxMDIFrame42s', None)
        AfxMDIFrame42s_2 = win32gui.FindWindowEx(AfxMDIFrame42s_1, None, 'AfxMDIFrame42s', None)
        AfxFrameOrView42 = win32gui.FindWindowEx(AfxMDIFrame42s_2, None, 'AfxFrameOrView42s', None)
        # win32gui.EnumChildWindows(handle, lambda hwnd, param: param.append(hwnd), hwndChildList)
        # for hwnd in hwndChildList:
        #     if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWi ndowVisible(hwnd):
        #         window_name = win32gui.GetClassName(hwnd)
        #         if window_name and str(window_name).find('AfxFrameOrView42s') >= 0:
        #             AfxFrameOrView42 = hwnd
        path = self.win.window_capture('GPS.bmp', AfxFrameOrView42)
        self.get_gps(path)
        msctls_statusbar32 = win32gui.FindWindowEx(self.radar_view_handle, None, 'msctls_statusbar32', None)
        path = self.win.window_capture('H.bmp', msctls_statusbar32)
        self.get_deepness(path)

    def get_gps(self, path):
        image = cv2.imread(path)
        image = image[80:120, 70:210]
        text = pytesseract.image_to_string(image, lang='chi_sim')
        print str(text).replace('\n', '').replace(' ', '')

    def get_deepness(self, path):
        image = cv2.imread(path)
        image = image[0:20, 150:250]
        # cv2.imshow('ss',image)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()
        text = pytesseract.image_to_string(image, lang='chi_sim')
        print str(text).replace('\n', '').replace(' ', '')

    # 2、获取句柄数据
    def get_handle(self, sub_class_name):
        master_win_endName = 'RadarView'
        self.desktop_handle = win32gui.GetDesktopWindow()
        if not self.radar_view_handle:
            self.radar_view_handle = self.handle_tool.find_subhandleForName(self.desktop_handle, master_win_endName)
            self.radar_view_name = win32gui.GetWindowText(self.radar_view_handle)
        return win32gui.FindWindowEx(self.radar_view_handle, None, sub_class_name, None)


# 句柄操作类
class Handle:

    def find_handle(self, winClass=None, winName=None):
        """
        找到一个窗口句柄
        :param winClass: 窗口类型
        :param winName: 窗口标题
        :return: 窗口句柄
        """
        return win32gui.FindWindow(winClass, winName)

    def __find_sub_handleAndName(self, handle, endwith):
        """
        根据父类窗口句柄查找子类窗口句柄和标题名称
        :param handle: 父类窗口句柄
        :return: 子类窗口和标题名
        """
        hwndChildList = []
        win32gui.EnumChildWindows(handle, lambda hwnd, param: param.append(hwnd), hwndChildList)

        for hwnd in hwndChildList:
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                window_name = win32gui.GetWindowText(hwnd)
                if window_name and str(window_name).find(endwith) >= 0:
                    return hwnd

    def __find_sub_handleAndName_(self, handle):
        """
        根据父类窗口句柄查找子类窗口句柄和标题名称
        :param handle: 父类窗口句柄
        :return: 子类窗口和标题名
        """
        hwndChildList = []
        win32gui.EnumChildWindows(handle, lambda hwnd, param: param.append(hwnd), hwndChildList)
        return hwndChildList

    @staticmethod
    def __get_windo_text(handle):
        """
        获取窗口标题
        :param handle: 窗口句柄
        :return: 窗口标题
        """
        return win32gui.GetWindowText(handle)

    @staticmethod
    def __get_windo_class(handle):
        return win32gui.GetClassName(handle)

    def find_subhandleForName(self, Phandle, endwith=None):
        """
        根据窗口标题名尾缀和父类句柄确定句柄
        :param Phandle: 父类句柄
        :param endwith: 窗口尾缀
        :return: 窗口句柄
        """
        sub_handle = self.__find_sub_handleAndName(Phandle, endwith)
        if not sub_handle:
            print '没有找到 %s 的句柄' % endwith
            return False
        return sub_handle

    def find_subhandleForClass(self, Phandle, classSrt=None):
        """
        根据窗口类名父类句柄确定句柄
        :param Phandle: 父类句柄
        :param endwith: 窗口尾缀
        :return: 窗口句柄
        """
        sub_list = self.__find_sub_handleAndName_(Phandle)
        sub_handles = []
        for sub_handle in sub_list:
            name = self.__get_windo_class(sub_handle)
            if str(name) == classSrt:
                print "%s <==> %x" % (name, sub_handle)
                sub_handles.append(sub_handle)
        return sub_handles

    def find_sub_handleForParentHandle(self, Phandle, className, name=None):
        """
        根据父句柄和类名查找窗口句柄
        :param Phandle: 父句柄
        :param className: 窗口类型
        :param name: 窗口标题
        :return: 窗口句柄
        """
        return win32gui.FindWindowEx(Phandle, 0, className, name)

    def GPS_setting(self, handleList, selectList):
        """
        设置GPS设置窗口的下拉列表
        :param handleList: 下拉句柄列表
        :param selectList: 下拉选项
        :return:无
        """
        for i in range(len(handleList)):
            self.__check_select_list(handleList[i], selectList[i])

    def __check_select_list(self, handle, No):
        """
        选择下拉列表
        :param handle: 列表句柄
        :param No: 第几个选项
        :return: 无
        """
        No_ = None
        if type(No) != int:
            No_ = int(No)
        else:
            No_ = No
        return win32gui.PostMessage(handle, win32con.CB_SETCURSEL, No_, 0)

    def send_message(self, handle, message):
        win32gui.SendMessage(handle, message, 0, 0)


class Windows:
    def get_imag_path(self, handle, name):
        """
        对指定窗口截图
        :param name: 图片名称
        :param handle: 指定窗口句柄
        :return:图片路径
        """
        im = ImageGrab.grab(win32gui.GetWindowRect(handle))
        # 参数 保存截图文件的路径
        im.save('..\\resources\\' + name)
        return str('..\\resources\\' + name)

    def window_capture(self, filename, hwnd=0):
        # 窗口的编号，0号表示当前活跃窗口
        #  根据窗口句柄获取窗口的设备上下文DC（Divice Context）
        hwndDC = win32gui.GetWindowDC(hwnd)
        # 根据窗口的DC获取mfcDC
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        # mfcDC创建可兼容的DC
        saveDC = mfcDC.CreateCompatibleDC()
        # 创建bigmap准备保存图片
        saveBitMap = win32ui.CreateBitmap()
        # 获取监控器信息
        MoniterDev = win32api.EnumDisplayMonitors(hwndDC, None)
        # 图片大小
        w = MoniterDev[0][2][2]
        h = MoniterDev[0][2][3]
        #  为bitmap开辟空间
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        #  高度saveDC，将截图保存到saveBitmap中
        saveDC.SelectObject(saveBitMap)
        #  截取从左上角（0，0）长宽为（w，h）的图片
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, '..\\resources\\' + filename)
        return '..\\resources\\' + filename

    def get_imag(self, handle):
        """
        对指定窗口截图
        :param name: 图片名称
        :param handle: 指定窗口句柄
        :return:图片
        """
        print win32gui.GetWindowRect(handle)
        return ImageGrab.grab(win32gui.GetWindowRect(handle))


class OpenCv:
    def contrast(self, source_image, dev_image, threshold=0.9):
        """
        识别GPS启用按钮是否已经勾选成功。
        :param source_image:元图片，界面截图
        :param dev_image:GPS启用按钮未使用截图
        :param threshold: 匹配紧缺度
        :return: True===》 未启用 需要操作 False===》启用 无需操作
        """
        # find the match position
        pos = ac.find_template(ac.imread(source_image), ac.imread(dev_image), threshold=threshold)
        if pos:
            # 有值
            return False
        return True
