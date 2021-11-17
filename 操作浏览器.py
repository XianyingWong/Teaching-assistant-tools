#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'这个文件包含了可以自动操作的代码，可以拿来引用'


from PIL import Image, ImageFont, ImageDraw, ImageGrab
import subprocess
import os
import aircv as ac
import win32api
import win32con
import time
import ctypes
import webbrowser
from pykeyboard import PyKeyboard
__all__ = [PyKeyboard]
from pymouse import PyMouseMeta
#__all__ = [PyMouse]
import cv2
import pyperclip


def 暂停(sleep_time):
    time.sleep(sleep_time)


def 黏贴(mystr):
    '''把文字（支持中文）复制黏贴'''
    pyperclip.copy(mystr)
    粘贴_不支持中文()

# trl v
def 粘贴_不支持中文():
    '''可用于黏贴'''
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(86, 0, 0, 0)
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    # print('ctrl_v')


def 单击(x, y):
    ''' 模拟电脑点击 '''
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    '''
    MOUSE_LEFTDOWN = 0x0002     # left button down 
    MOUSE_LEFTUP = 0x0004       # left button up 
    MOUSE_RIGHTDOWN = 0x0008    # right button down 
    MOUSE_RIGHTUP = 0x0010 十六进制的10，即十进制的16     # right button up 
    MOUSE_MIDDLEDOWN = 0x0020   # middle button down 
    MOUSE_MIDDLEUP = 0x0040     # middle button up 
    '''


def 右键():
    ''' 模拟电脑点击右键 '''
    # ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(8, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(16, 0, 0, 0, 0)


# 双击
def 双击(x, y):
    ''' 模拟电脑双击 '''
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)

# 关闭浏览器
def ctrl_w():
    '''ctrl + w，可用于关闭浏览器'''
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(87, 0, 0, 0)
    win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

# 全选
def 全选():
    '''ctrl + a，可用于全选'''
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(65, 0, 0, 0)
    win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    # print('ctrl_w')

# trl c
def 复制():
    '''ctrl + c，可用于复制'''
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(67, 0, 0, 0)
    win32api.keybd_event(67, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    # print('ctrl_c')


# enter
def 回车():
    '''回车'''
    win32api.keybd_event(0x0D, 0, 0, 0)
    win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes

def page_down():
    '''翻页'''
    ctypes.windll.user32.keybd_event(34, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    ctypes.windll.user32.keybd_event(34, 0, win32con.KEYEVENTF_KEYUP, 0)

# 匹配图片
def 匹配图片(myScreencap, mypng):
    '''匹配图片, 返回True或者False'''
    if _matchImg(myScreencap, mypng) is not None:
        print("匹配图片" + mypng + str(
            _matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            _matchImg(myScreencap, mypng)['result'][1]))
        myx = str(_matchImg(myScreencap, mypng)['result'][0])
        myy = str(_matchImg(myScreencap, mypng)['result'][1])
        # click(int(float(myx)), int(float(myy)))
        # print("-------------对比图片return True")
        return True
        # time.sleep(2)
    else:
        # print("-------------对比图片return False")
        return False

# my_img_height为小图片的高度，即为每次增加的像素,148
# 1080* 2280
def 匹配图片_从上到下(imgsrc, imgobj, my_img_height):
    '''从顶部朝底部寻找匹配，点击第一个'''
    img = cv2.imread(imgsrc)
    y0, y1, x0, x1 = 0, 0, 0, 1920
    for step_y1 in range(my_img_height, 1080, my_img_height):
        cropped = img[y0:step_y1, x0:x1]  # 裁剪坐标为[y0:y1, x0:x1]
        cv2.imwrite('tmp_' + imgsrc, cropped)
        if 匹配图片('tmp_' + imgsrc, imgobj):
            # print('找到了，' + str(step_y1))
            if _matchImg('tmp_' + imgsrc, imgobj) is not None:
                # myx = str(matchImg('tmp_' + imgsrc, imgobj)['result'][0])
                # myy = str(matchImg('tmp_' + imgsrc, imgobj)['result'][1])
                # os.popen('adb -s QMFDU20630010700 shell input tap ' + myx + ' ' + myy, 'r', 1)
                # time.sleep(4)
                myx = str(_matchImg('tmp_' + imgsrc, imgobj)['result'][0])
                myy = str(_matchImg('tmp_' + imgsrc, imgobj)['result'][1])
                myx_off = str(_matchImg('tmp_' + imgsrc, imgobj)
                              ['result'][0] - 0)
                # os.popen('adb -s QMFDU20630010700 shell input tap ' + myx_off + ' ' + myy, 'r', 1)
                print(myx_off, myy)
                单击(int(float(myx_off)), int(float(myy)))
                time.sleep(0.5)
            break
        else:
            pass

# 打开浏览器
# webbrowser.open(url, new=0, autoraise=True)
# 在系统的默认浏览器中访问url地址，
# 如果new = 0, url会在同一个 浏览器窗口中打开；
# 如果new = 1，新的浏览器窗口会被打开;
# 如果new = 2 新的浏览器tab会被打开
def 打开网页(url):
    '''打开浏览器'''
    # print('    # 打开' + url)
    webbrowser.open(url, new=0)


# 模拟键盘输入字符串
def 输入(my_string):
    '''模拟键盘输入字符串'''
    # 定义鼠标键盘实例
    k = PyKeyboard()
    k.type_string(my_string)


# 对比图片并点击
def 匹配图片单击(myScreencap, mypng):
    '''对比图片并点击'''
    if _matchImg(myScreencap, mypng) is not None:
        print("-------------点击按钮！" + mypng + str(
            _matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            _matchImg(myScreencap, mypng)['result'][1]))
        myx = str(_matchImg(myScreencap, mypng)['result'][0])
        myy = str(_matchImg(myScreencap, mypng)['result'][1])
        单击(int(float(myx)), int(float(myy)))
        # time.sleep(3)
        print("-------------结束点击按钮。")


# 等待对比图片并点击
def 匹配图片等待单击(mypng):
    '''等待对比图片并点击，一出现就点击'''
    find_it = False
    while find_it == False:
        截图('myScreencap.png')
        time.sleep(1)
        if _matchImg('myScreencap.png', mypng) is not None:
            print("-------------点击按钮！" + mypng + str(
                _matchImg('myScreencap.png', mypng)['result'][0]) + ',' + str(
                _matchImg('myScreencap.png', mypng)['result'][1]))
            myx = str(_matchImg('myScreencap.png', mypng)['result'][0])
            myy = str(_matchImg('myScreencap.png', mypng)['result'][1])
            单击(int(float(myx)), int(float(myy)))
            # time.sleep(3)
            print("-------------结束点击按钮。")
            find_it = True
        else:
            find_it = False


# 对比图片并点击
def 匹配图片右键单击(myScreencap, mypng):
    '''对比图片并点击'''
    if _matchImg(myScreencap, mypng) is not None:
        print("-------------点击按钮！" + mypng + str(
            _matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            _matchImg(myScreencap, mypng)['result'][1]))
        myx = str(_matchImg(myScreencap, mypng)['result'][0])
        myy = str(_matchImg(myScreencap, mypng)['result'][1])
        右键(int(float(myx)), int(float(myy)))
        # time.sleep(3)
        print("-------------结束点击按钮。")


# 截图
def 截图(im_name):
    '''截图 ：参数要加后缀，比如.png或者.jpg'''
    im = ImageGrab.grab()
    # 放到pic文件夹下
    im.save(im_name)


# 组合按键
def 组合按键(key1, key2):
    '''模拟电脑键盘按组合按键: https://blog.csdn.net/zhanglidn013/article/details/35988381 或者 https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    '''
    ctypes.windll.user32.keybd_event(key1, 0, 0, 0)
    ctypes.windll.user32.keybd_event(key2, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    ctypes.windll.user32.keybd_event(key2, 0, win32con.KEYEVENTF_KEYUP, 0)
    ctypes.windll.user32.keybd_event(key1, 0, win32con.KEYEVENTF_KEYUP, 0)


# 按下一个按钮
def 按下键盘(key1):
    '''模拟电脑键盘按下一个按键'''
    ctypes.windll.user32.keybd_event(key1, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    ctypes.windll.user32.keybd_event(key1, 0, win32con.KEYEVENTF_KEYUP, 0)
    '''
    A　　　65
    B　　　66
    C　　　67
    D　　　68
    E　　　69
    F　　　70
    G　　　71
    H　　　72
    I　　　73
    J　　　74
    K　　　75
    L　　　76
    M　　　77
    N　　　78
    O　　　79
    P　　　80
    Q　　　81
    R　　　82
    S　　　83
    T　　　84
    U　　　85
    V　　　86
    W　　　8
    X　　　88
    Y　　　89
    Z　　　90
    0　　　48
    1　　　49
    2　　　50
    3　　　51
    4　　　52
    5　　　53
    6　　　54
    7　　　55
    8　　　56
    9　　　57

    小键盘
    0 　　96 
    1　　 97 
    2 　　98 
    3　　 99 　
    4 　　100
    5 　　101
    6　　 102
    7 　　103
    8 　　104　
    9　　 105　
    * 　　106
    + 　　107
    Enter 108　
    -　　 109
    . 　　110
    / 　　111

    F1 　　112 
    F2 　　113　
    F3 　　114　
    F4 　　115　　
    F5 　　116　
    F6 　　117　
    F7 　　118 
    F8 　　119　　
    F9 　　120　　
    F10　　121　　
    F11　　122　
    F12　　123　

    Backspace 　　　8 
    Tab 　　　　　　9 
    Clear 　　　　12 
    Enter 　　　　　13 
    Shift　　　　　 16 
    Control 　　　　17 
    Alt 　　　　　　18
    Caps Lock 　　　20 
    Esc 　　　　　　27 
    Spacebar　　　　32 
    Page Up　　　　 33 
    Page Down 　　　34
    End 　　　　　　35
    Home　　　　　36 
    Left Arrow　　　37 
    Up Arrow　　　　38 
    Right Arrow 　　39 
    Down Arrow 　　 40
    Insert 　　　　 45 
    Delete 　　　　 46 
    Help 　　　　　 47 
    Num Lock 　　　 144
    '''


def 按下键盘2(key1):
    '''模拟电脑键盘按下一个按键（第二种方法）'''
    # win32api.keybd_event(0x0D, 0, 0, 0)
    # win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(key1, 0, 0, 0)
    win32api.keybd_event(key1, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟鼠标滑动操作：线性
def 鼠标滑动(x, y, x1, y1):
    '''模拟鼠标滑动操作：线性'''
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(1)
    ctypes.windll.user32.SetCursorPos(x1, y1)
    time.sleep(1)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)


# 模拟人鼠标滑动操作：缓动滑动（非线性）
def 鼠标缓慢滑动(x, y, x_1, y_1):
    m = PyMouseMeta()
    my_x_1_half = (x_1 - x)
    # my_x_2 = x + my_x_1_half
    # my_x_3 = x + my_x_1_half + my_x_1_half
    # print(math.ceil(float(my_x_1_half)))
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(1)
    track_list = _get_track(my_x_1_half)
    for track in track_list:
        x = x + track
        # print('for: track is ' + str(x))
        m.move(int(x), y_1)
        time.sleep(0.02)
    # m.move(int(my_x_2), y_1)
    # time.sleep(1)
    # m.move(int(my_x_3), y_1)
    # time.sleep(2)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)


def 按日期排序文件(file_path):
    files = os.listdir(file_path)
    if not files:
        return
    else:
        files = sorted(files, key=lambda x: os.path.getmtime(
            os.path.join(file_path, x)))  # 格式解释:对files进行排序.x是files的元素,:后面的是排序的依据.   x只是文件名,所以要带上join.
        return files


def 删除文件(path_data):
    # 判断文件是否存在
    # if os.path.exists('./sound/filelist.txt'):
    #     os.remove('./sound/filelist.txt')
    # else:
    #     print("要删除的文件不存在！a+ 去生成")
    for i in os.listdir(path_data):  # os.listdir(path_data)#返回一个列表，里面是当前目录下面的所有东西的相对路径
        file_data = path_data + "\\" + i  # 当前文件夹的下面的所有东西的绝对路径
        # os.path.isfile判断是否为文件,如果是文件,就删除.如果是文件夹.递归给del_file.
        if os.path.isfile(file_data):
            os.remove(file_data)
        else:
            删除文件(file_data)


# 图片添加字
def _add_num(im01, mypng, x, y):
    img = Image.open(im01)
    imgmypng = Image.open(mypng)
    # ImageDraw.Draw()函数会创建一个用来对image进行操作的对象，
    # 对所有即将使用ImageDraw中操作的图片都要先进行这个对象的创建。
    draw = ImageDraw.Draw(img)

    # 设置字体和字号‪C:\Windows\Fonts\msyhbd.ttc
    myfont = ImageFont.truetype('C:/windows/fonts/msyhbd.ttc', size=80)

    # 设置要添加的数字的颜色为红色
    fillcolor = "#2F83C7"

    # 昨天博客中提到过的获取图片的属性
    width, height = imgmypng.size

    # 设置添加数字的位置，具体参数可以自己设置，从左上角开始
    draw.text((x - width / 2, y - height / 2),
              '99999999', font=myfont, fill=fillcolor)

    # 将修改后的图片以格式存储
    img.save(im01, 'png')

    return 0


# 滑动鼠标内部用到的函数
def _get_track(distance):
    # 拿到移动轨迹，模仿人的滑动行为，先匀加速后匀减速
    # 匀变速运动基本公式：
    # ①v=v0+at
    # ②s=v0t+(1/2)at²
    # ③v²-v0²=2as
    #
    # :param distance: 需要移动的距离
    # :return: 存放每0.2秒移动的距离
    # 初速度
    v = 0
    # 单位时间为0.2s来统计轨迹，轨迹即0.2内的位移
    t = 0.1
    # 位移/轨迹列表，列表内的一个元素代表0.2s的位移
    tracks = []
    # 当前的位移
    current = 0
    # 到达mid值开始减速
    mid = distance * 4 / 5
    distance += 10  # 先滑过一点，最后再反着滑动回来
    while current < distance:
        if current < mid:
            # 加速度越小，单位时间的位移越小,模拟的轨迹就越多越详细
            a = 2  # 加速运动
        else:
            a = -3  # 减速运动
        # 初速度
        v0 = v
        # 0.2秒时间内的位移
        s = v0 * t + 0.5 * a * (t ** 2)
        # 当前的位置
        current += s
        # 添加到轨迹列表
        tracks.append(round(s))
        # 速度已经达到v,该速度作为下次的初速度
        v = v0 + a * t
    # 反着滑动到大概准确位置
    for i in range(3):
        tracks.append(-2)
    for i in range(4):
        tracks.append(-1)
    return tracks


# 对比两张图，找到坐标。
def _matchImg(imgsrc, imgobj):  # imgsrc=原始图像，imgobj=待查找的图片
    imsrc = ac.imread(imgsrc)
    imobj = ac.imread(imgobj)
    match_result = ac.find_template(imsrc, imobj, 0.8)
    # 0.9、confidence是精度，越小对比的精度就越低 {'confidence': 0.5435812473297119,
    # 'rectangle': ((394, 384), (394, 416), (450, 384), (450, 416)), 'result': (422.0, 400.alipay_leave0)}
    if match_result is not None:
        match_result['shape'] = (imsrc.shape[1], imsrc.shape[0])  # 0为高，1为宽
    return match_result

