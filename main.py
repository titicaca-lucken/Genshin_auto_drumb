import pythoncom,win32api,win32con,time,win32gui
import cv2
import os,sys
import numpy as np
from PIL import ImageGrab
import threading,multiprocessing

# cv2.namedWindow("img", 0)
# cv2.resizeWindow("img",256,150)
def image_match(yuan,icon,d):
    w,h=d,d
    # d在进行匹配测试调节锚框大小时可用,其他时段为无效参数
    img_=yuan.copy()
    methods=['cv2.TM_CCOEFF','cv2.TM_CCOEFF_NORMED','cv2.TM_CCORR','cv2.TM_SQDIFF','cv2.TM_SQDIFF_NORMED']
    # 个人测试效果ccoeff>corr=sqdiff
    method=eval('cv2.TM_CCOEFF_NORMED')
    res=cv2.matchTemplate(img_,icon,method)
    min_val,max_val,min_loc,max_loc=cv2.minMaxLoc(res)
    top_left=max_loc
    # bottom_right=(top_left[0]+w,top_left[1]+h)
    # cv2.rectangle(img_,top_left,bottom_right,(0,0,255),2)
    # cv2.namedWindow("img", 0)
    # cv2.resizeWindow("img",128,74)
    # cv2.imshow('img',img_)
    # time.sleep(1)
    # cv2.waitKey(3)

    return max_val,top_left

def click(im0,icon1,icon2,loc_key,key_ascii,gamma):
    top_left=(loc_key[0]+100,loc_key[1]-60)

    bottom_right=(top_left[0]+160,top_left[1]+160)
    count=0
    mini_count=0
    # color=(0,255,0)
    for i in range(3100):
        # im0 =ImageGrab.grab((x1,y1,x2,y2))
        im0=ImageGrab.grab((top_left[0], top_left[1], bottom_right[0], bottom_right[1]))
        img = cv2.cvtColor(np.array(im0), cv2.COLOR_RGB2BGR) 
        v1,_=image_match(img,icon1,80)
        v2,_=image_match(img,icon2,80)

        # print('v1',v1)
        # print('v2',v2)
        if float(v1)>0.8*gamma:
            # print('v1',v1)
            count=0
            mini_count=0
            win32api.keybd_event(key_ascii,0,win32con.KEYEVENTF_KEYUP,0)
            win32api.keybd_event(key_ascii,0,0,0)
            win32api.keybd_event(key_ascii,0,win32con.KEYEVENTF_KEYUP,0)
            continue
        if count>0:
            if float(v2)<0.7*gamma*gamma and mini_count>5:
                win32api.keybd_event(key_ascii,0,0,0)
                win32api.keybd_event(key_ascii,0,win32con.KEYEVENTF_KEYUP,0)
                count=0
                mini_count=0
            else:
                mini_count+=1
                win32api.keybd_event(key_ascii,0,0,0)
                # win32api.keybd_event(75,0,win32con.KEYEVENTF_KEYUP,0)
        else:
            if float(v2)>0.75*gamma:
                # print('v2',v2)
                win32api.keybd_event(key_ascii,0,0,0)
                count+=1

def main():
    py_path = sys.argv[0][:-7]
    # print(py_path)
    icon_a = cv2.imread(py_path+"/icon/%s.jpg"%('a'))
    icon_s = cv2.imread(py_path+"/icon/%s.jpg"%('s'))
    icon_d = cv2.imread(py_path+"/icon/%s.jpg"%('d'))
    icon_j = cv2.imread(py_path+"/icon/%s.jpg"%('j'))
    icon_k = cv2.imread(py_path+"/icon/%s.jpg"%('k'))
    icon_l = cv2.imread(py_path+"/icon/%s.jpg"%('l'))
    icon1 = cv2.imread(py_path+"/icon/%s.jpg"%('icon_orange'))
    icon2 = cv2.imread(py_path+"/icon/%s.jpg"%('icon_purple'))
    keys=[65,83,68,74,75,76]
    wdname = '原神'

    hwnd = win32gui.FindWindow(None, wdname)
    win32gui.SetForegroundWindow(hwnd)
    # 使窗体最大化
    # win32gui.ShowWindow(hwnd,win32con.SW_MAXIMIZE)

    x1,y1,x2,y2=win32gui.GetWindowRect(hwnd)
    # print(x1,y1,x2,y2)

    x1,y1,x2,y2=int(x1*1.25)+9,int(y1*1.25),int(x2*1.25)-7,int(y2*1.25)-8  
    # 此处的1.25,是与win10"缩放与布局"功能中文本放大比例125%相匹配,其他比例需要自行调节
    # win涉及窗口缩放的功能都很坑
    
    time.sleep(5) # 等待5s以切换到游戏窗口
    im0 =ImageGrab.grab((x1,y1,x2,y2))
    img = cv2.cvtColor(np.array(im0), cv2.COLOR_RGB2BGR) 
    img = np.array(img).astype('uint8')
    _,loc_a=image_match(img,icon_a,130)
    _,loc_s=image_match(img,icon_s,130)
    _,loc_d=image_match(img,icon_d,130)
    _,loc_j=image_match(img,icon_j,130)
    _,loc_k=image_match(img,icon_k,130)
    _,loc_l=image_match(img,icon_l,130)

    thread_list = []

    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_a,keys[0],0.7,)))
    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_s,keys[1],0.8,)))
    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_d,keys[2],0.7,)))
    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_j,keys[3],0.7,)))
    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_k,keys[4],0.8,)))
    thread_list.append(multiprocessing.Process(target=click,args=(im0,icon1,icon2,loc_l,keys[5],0.7,)))

    for t in thread_list:
        t.start()
        # shibaxiancheng.q[i]
    # for t in thread_list:
    #     t.setDaemon(True)  # 设置为守护线程，不会因主线程结束而中断
    #     t.start()

    for t in thread_list:
        t.join()  # 子线程全部加入，主线程等所有子线程运行完毕

if __name__ == "__main__":
    main()