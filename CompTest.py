#coding=gbk
import win32com.client
from win32api import GetSystemMetrics
import win32gui
from ctypes import windll
import time
import hashlib
import os
import _winreg
import win32api
import subprocess
import types

g_tipsTitle = 'gj - tips'

def IsVisibleWnd(hwnd, wndLst):
    if(win32gui.IsWindowVisible(hwnd)):
        wndLst.append(hwnd)

def IsWindowOpened(title):
    wndLst = []
    win32gui.EnumWindows(IsVisibleWnd, wndLst)
    for hwnd in wndLst:
        wndText = win32gui.GetWindowText(hwnd)

        if (wndText.find(title) != -1):
            return True
    return False

def GetWindowHandle(title):
    wndLst = []
    win32gui.EnumWindows(IsVisibleWnd, wndLst)
    for hwnd in wndLst:
        wndText = win32gui.GetWindowText(hwnd)

        if (wndText.find(title) != -1):
            hwnd
    return 0

def EnumVisibleWndProc(_hwnd, _lParam):
    '''
    @desc: 枚举桌面可视窗口
    '''
    if(win32gui.IsWindowVisible(_hwnd)):
        isWndVisibleList.append(_hwnd)
def FindAllVisibleWnd():
    '''
    @desc: 返回系统中的所有可见的窗口
    '''
    global isWndVisibleList
    isWndVisibleList = []
    win32gui.EnumWindows(EnumVisibleWndProc, 0)
    return isWndVisibleList

def FindWindow(_className, _title, _fullMatch=False):
    '''
@desc: 根据窗口类名和标题查找窗口,以数组形式返回符合类名的所有窗口    
@param:_className  窗口的类名
@param:_title  窗口标题
@param:_fullMatch  是否执行完全匹配查找，默认为False
异常原因：1.参数不正确
''' 
    if type(_title) == types.UnicodeType:
        print _title
        _title = _title.encode('gbk')
    wndList = []
    for hwnd in FindAllVisibleWnd():
        windowText = win32gui.GetWindowText(hwnd)
        if type(windowText) == types.UnicodeType:
            windowText = windowText.encode('gbk')
         
        if _fullMatch:
            if(cmp(windowText, _title) == 0): 
                if(cmp(win32gui.GetClassName(hwnd), _className) == 0):
                    wndList.append(hwnd)
        else:
            if(windowText.upper().find(_title.upper()) >= 0): 
                if(win32gui.GetClassName(hwnd).upper().find(_className.upper()) >= 0):                    
                    wndList.append(hwnd)
    return wndList

def IsProcessExist(processName):
    WMI = win32com.client.GetObject('winmgmts:') 
    processCodeCov = WMI.ExecQuery('select * from Win32_Process where Name="%s"' % processName) 
    return (len(processCodeCov) > 0)

def GetMd5(filePath):
    if os.path.exists(filePath):
        md5 = hashlib.md5()
        fp = open(filePath, "rb")
        blockSize = 1024
        bytesBlock = fp.read(blockSize)
        while bytesBlock != b'':
            md5.update(bytesBlock)
            bytesBlock = fp.read(blockSize)
        fp.close()
        return md5.hexdigest().lower()

def TakeScreenShot():
    #print 'grab image'
    imageName = '%s.png' % time.strftime('%Y%m%d%H%M%S')
    picPath = './images/' + imageName
    #
    print 'imageName = %s, pic_path = %s' % (imageName, picPath)
    dll = windll.LoadLibrary('CCommonFunDll.dll')
    width = GetSystemMetrics(0)
    height = GetSystemMetrics (1)
    dll.CF_CutScreenToFile(0, 0, width, height, picPath)
    time.sleep(2)

def WebsitesTest(websites):
    passed = 0
    failed = 0
    for website in websites:
        try:
            ie = win32com.client.DispatchEx('InternetExplorer.Application.1')
            ie.Visible = 1    
            ie.Navigate(website)
            while ie.Busy:
                time.sleep(1)

            time.sleep(5) # stay for 5 seconds for tips to show up or crash
            
            if IsWindowOpened(g_tipsTitle) == True:
                failed += 1
                print website + " test fail"
                TakeScreenShot()
            else:
                passed += 1
                print website + " test pass"


            ie.quit()  # if crash, Exception will be throw here
            time.sleep(3)
        except Exception, ex:
            failed += 1
            print website + " test fail"
            TakeScreenShot()
            print Exception, ":", ex

    print "Passed: %d; Failed: %d" % (passed, failed)


def WinwordTest(fileNameToOpen, fileNameToSave):
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = 1
        app.DisplayAlerts = 0
        doc = app.Documents.Open(fileNameToOpen)
        time.sleep(3)

        if os.path.exists(fileNameToSave + '.txt'):            
            os.remove(fileNameToSave + '.txt')        
        doc.SaveAs(fileNameToSave, 4)

        time.sleep(3)


        mystring = ""
        mystring = open(fileNameToSave + '.txt', 'r').read()
        #ret = cmp('TeamA3Automation\n', mystring);
        if mystring.find("TeamA3Automation") != -1 :
            print "Word test pass"
        else:
            print "Word test fail"
            TakeScreenShot()
        
        app.ActiveDocument.Close(SaveChanges=True)
        app.Quit()

    except Exception, e:
            print e
            print "Word test fail"
            
            TakeScreenShot()

def ExcelTest(fileNameToOpen, fileNameToSave):
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = 1
        app.DisplayAlerts = 0
        doc = app.Workbooks.Open(fileNameToOpen)
        time.sleep(3)

        if os.path.exists(fileNameToSave + '.csv'):
            os.remove(fileNameToSave + '.csv')        
        doc.SaveAs(fileNameToSave, 6)
        time.sleep(3)


        mystring = ""
        mystring = open(fileNameToSave + '.csv', 'r').read()
        #ret = cmp('TeamA3Automation\n', mystring);
        if mystring.find("TeamA3Automation") != -1 :
            print "Excel test pass"
        else:
            print "Excel test fail"
            TakeScreenShot()
        
        doc.Close(SaveChanges=True)
        app.Quit()

    except Exception, e:
            print e
            print "Excel test fail" 
            TakeScreenShot()


def PowerpointTest(fileNameToOpen, fileNameToSave):
    try:
        app = win32com.client.DispatchEx("Powerpoint.Application")
        app.Visible = 1
        app.DisplayAlerts = 0
        doc = app.Presentations.Open(fileNameToOpen)

        time.sleep(3)

        if os.path.exists(fileNameToSave + '.htm'):
            os.remove(fileNameToSave + '.htm')        
        doc.SaveAs(fileNameToSave, 12)
        time.sleep(3)


        mystring = ""
        mystring = open(fileNameToSave + '.htm', 'r').read()
        #ret = cmp('TeamA3Automation\n', mystring);
        if mystring.find("TeamA3Automation") != -1 :
            print "Powerpoint test pass"
        else:
            print "Powerpoint test fail"
            TakeScreenShot()
        
        doc.Close()
        app.Quit()

    except Exception, e:
            print e
            print "Powerpoint test fail" 
            TakeScreenShot()


def OutlookTest(fileNameToSave):
    try:
        app = win32com.client.gencache.EnsureDispatch('Outlook.Application')
        
        time.sleep(3)

        mail = app.CreateItem(0)

        recip = mail.Recipients.Add('yerikli@tencent.com')
        subj = mail.Subject = 'TeamA3Automation'

        body = ["TeamA3Automation"]
        body.append('TeamA3Automation')
        mail.Body = '\r\n'.join(body)
        
        if os.path.exists(fileNameToSave + '.htm'):
            os.remove(fileNameToSave + '.htm')
            
        mail.SaveAs(fileNameToSave + '.htm', win32com.client.constants.olHTML)
        time.sleep(3)

        mystring = ""
        mystring = open(fileNameToSave + '.htm', 'r').read()
        #ret = cmp('TeamA3Automation\n', mystring);
        if mystring.find("TeamA3Automation") != -1 :
            print "Outlook test pass"
        else:
            print "Outlook test fail"
            TakeScreenShot()
        

        app.Quit()


    except Exception, e:
            print e
            print "Outlook test fail" 
            TakeScreenShot()


def RunSoft(Dir, Caption, Class, exeName, param, bFullMatch, waitTime=60): 
    # 开始记录时间
    Dir = Dir + " " + param 
    beforeTime = win32api.GetTickCount() 

    if IsProcessExist(exeName):
        strKillRC = "taskKill /F /IM " + exeName
        os.system(strKillRC)      
        time.sleep(1) 
    
    subprocess.Popen('%s' % Dir, shell=True,) 
    
    bRes = False
    
    while(True):
        res = win32api.GetTickCount() - beforeTime 
        if res > waitTime * 1000 :
            break
        wndList = FindWindow(Class, Caption, bFullMatch) 
        if len(wndList) > 0 :  
            # 等几秒后截图
            time.sleep(10)  # 等待10秒，如果没有tips，且窗口还在就算pass

            if IsWindowOpened(g_tipsTitle) == True:
                print exeName + " test fail"
                TakeScreenShot()
                return False

            wndList1 = FindWindow(Class, Caption, bFullMatch) 
            if len(wndList1) > 0:
                strKillRC = "taskKill /F /IM " + exeName
                os.system(strKillRC)      
                time.sleep(1)
            else:
                print exeName + " test fail"
                TakeScreenShot()
                return False
        
            
            bRes = True
            break
        time.sleep(1)   
    return bRes


def AcrobatRdTest(fileToOpen, shortFileName):
    regPath = r"SOFTWARE\Adobe\Acrobat Reader\11.0\Installer"
    regKey = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, regPath)
    value,type = _winreg.QueryValueEx(regKey, "Path")

    Dir  = '"' + value + r'Reader\AcroRd32.exe"'
    Caption = shortFileName + r" - Adobe Reader"
    Class = "AcrobatSDIWindow"
    exeName = "AcroRd32.exe"
    param = fileToOpen
    bFullMatch = True
    bRes = RunSoft(Dir, Caption, Class, exeName, param, bFullMatch) 
    if bRes == True :
        print "Acrobat Reader test pass"
    else:
        print "Acrobat Reader test fail"
        TakeScreenShot()
    return bRes


def WpsTest(fileNameToOpen, fileNameToSave):
    try:
        app = win32com.client.DispatchEx("KWPS.Application")
        app.Visible = 1
        app.DisplayAlerts = 0
        doc = app.Documents.Open(fileNameToOpen)
        time.sleep(3)

        if os.path.exists(fileNameToSave):            
            os.remove(fileNameToSave)        
        doc.SaveAs(fileNameToSave,  2)

        time.sleep(3)


        mystring = ""
        mystring = open(fileNameToSave, 'r').read()
        #ret = cmp('TeamA3Automation\n', mystring);
        if mystring.find("TeamA3Automation") != -1 :
            print "Wps test pass"
        else:
            print "Wps test fail"
            TakeScreenShot()
        
        app.ActiveDocument.Close(SaveChanges=True)
        app.Quit()

    except Exception, e:
            print e
            print "Wps test fail"
            
            TakeScreenShot()






def WMPlayerTest(fileToOpen):


    Dir  = r'"C:\Program Files\Windows Media Player\wmplayer.exe"'
    Caption = "Windows Media Player"
    Class = "CabinetWClass"
    exeName = "wmplayer.exe"
    param = fileToOpen
    bFullMatch = True
    bRes = RunSoft(Dir, Caption, Class, exeName, param, bFullMatch) 
    if bRes == True :
        print "Windows Media Player test pass"
    else:
        print "Windows Media Player test fail"
        TakeScreenShot()
    return bRes

def Rundll32Test():
    Dir  = r'rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1'
    Caption = "添加或删除程序"
    Class = "NativeHWNDHost"
    exeName = "rundll32.exe"
    param = ""
    bFullMatch = True
    bRes = RunSoft(Dir, Caption, Class, exeName, param, bFullMatch) 
    if bRes == True :
        print "rundll32.exe test pass"
    else:
        print "rundll32.exe test fail"
        TakeScreenShot()
    return bRes

# Main Logic from here




websites = [
    'www.alipay.com', 'https://my.alipay.com/portal/i.htm?referer=https%3A%2F%2Fauth.alipay.com%2Flogin%2FhomeB.htm%3FredirectType%3Dparent',
    'www.baifubao.com', 'https://www.baifubao.com/user/0/login/0?ru=%2Fuser%2F0%2Fmy_bfb%2F0', 'http://life.baifubao.com/transfer/0/start/0', 'http://life.baifubao.com/interbank/0/start/0',
    'www.tenpay.com', 'https://www.tenpay.com/v2/index.shtml?target=%2Fapp%2Fv1.0%2Fcftaccount.cgi%3FADTAG%3DTENPAY_V2.CFTACCOUNT.SIDERBAR_TOP.MYACCOUNT',
    'www.qq.com', 'vip.qq.com', 'web.qq.com', 'web2.qq.com', 'pengyou.qq.com', 'mail.qq.com', 'qzone.qq.com', 'http://t.qq.com/', 'dnf.qq.com', 
    'www.tencent.com', 'www.soso.com', 'www.paipai.com', 'http://member.paipai.com/cgi-bin/login_entry?PTAG=20257.1.7',        
    'weibo.com', 'http://mail.163.com/', 'http://mail.sina.com.cn/',
    'http://www.renren.com/', 'http://www.youku.com/', 'http://www.tudou.com/',
    'http://www.etao.com', 'http://login.etao.com/?spm=1002.1.1.3.D3qY8C&redirect_url=http%3A%2F%2Fwww.etao.com%2F%3Ftbpm%3D20140704&logintype=taobao'
    'www.taobao.com', 'https://login.taobao.com/member/login.jhtml?spm=1.7274553.1997563269.1.nrPFB7&f=top&redirectURL=http%3A%2F%2Fwww.taobao.com%2F',
    'http://www.baidu.com', 'http://www.sina.com.cn', 'www.sohu.com', 'www.163.com', 'www.iqiyi.com', 'www.ifeng.com',
    'www.4399.com', 'www.17173.com', 'www.xinhuanet.com',
    'www.jd.com', 'https://passport.jd.com/new/login.aspx?ReturnUrl=http%3A%2F%2Fwww.jd.com%2F', 'http://chat9.jd.com/index.action?pid=1127691&price=&stock=%25E5%258C%2597%25E4%25BA%25AC%25E6%259C%259D%25E9%2598%25B3%25E5%258C%25BA%25E7%25AE%25A1%25E5%25BA%2584%25EF%25BC%2588%25E6%259C%2589%25E8%25B4%25A7%25EF%25BC%2589&score=5&commentNum=649&imgUrl=g16%252FM00%252F0B%252F13%252FrBEbRVN9qX8IAAAAAAH0zQiEQgsAACQpAO6cfUAAfTl578.jpg&wname=%25E8%2581%2594%25E6%2583%25B3%25EF%25BC%2588Lenovo%25EF%25BC%2589%2520Y430p%252014.0%25E8%258B%25B1%25E5%25AF%25B8%25E7%25AC%2594%25E8%25AE%25B0%25E6%259C%25AC%25E7%2594%25B5%25E8%2584%2591%25EF%25BC%2588i7-4710MQ%25208G%25201T%2520GTX850M%25202G%25E7%258B%25AC%25E6%2598%25BE%2520DVD%25E5%2588%25BB%2520%25E6%2591%2584%25E5%2583%258F%25E5%25A4%25B4%2520Win8%25EF%25BC%2589%25E9%25BB%2591%25E8%2589%25B2&advertiseWord=%25E3%2580%2590%25E6%2596%25B0Y%25E7%25A5%259E%25E3%2581%25AE%25E6%2580%25A7%25E8%2583%25BD%25E5%25BC%25BA%25E5%2588%25B0%25E6%25B2%25A1%25E6%259C%258B%25E5%258F%258B%25EF%25BC%2581%25E3%2580%2591%255C(%255Eo%255E)%252F~%25E6%2583%25B9%25E6%2588%2591%25E5%2585%25B6%25E8%25B0%2581%25CE%25A9%25E5%258C%2597%25E9%25A3%258E%25E4%25B9%258B%25E7%25A5%259E%25E6%2588%2598%25E7%2595%25A5%25E7%25BA%25A7%2520%25E2%2588%2591&seller=%25E8%2581%2594%25E6%2583%25B3%25EF%25BC%2588Lenovo%25EF%25BC%2589&evaluationRate=&recent=&code=1&area=%25E5%258C%2597%25E4%25BA%25AC%25E6%259C%259D%25E9%2598%25B3%25E5%258C%25BA%25E7%25AE%25A1%25E5%25BA%2584&size=&services=%25E7%2594%25B1%2520%25E4%25BA%25AC%25E4%25B8%259C%2520%25E5%258F%2591%25E8%25B4%25A7%25E5%25B9%25B6%25E6%258F%2590%25E4%25BE%259B%25E5%2594%25AE%25E5%2590%258E%25E6%259C%258D%25E5%258A%25A1%25E3%2580%2582',
    'www.tmall.com', 'http://login.tmall.com/?spm=3.7328325.a2226mz.1.aUWIyB&redirect_url=http%3A%2F%2Fwww.tmall.com%2F',
    'www.ctrip.com', 'https://accounts.ctrip.com/member/login.aspx?BackUrl=http%3A%2F%2Fwww.ctrip.com%2F&responsemethod=get',
    'www.12306.com', 'www.easymoney.com',
    'www.58.com', 'https://passport.58.com/login',
    'www.soufun.com', 'http://passport.soufun.com/',
    'www.jumei.com', 'http://www.jumei.com/i/account/login', 
    'www.liepin.com', 'www.jiayuan.com',
    'www.ganji.com', 'https://passport.ganji.com/login.php?next=/', 
    'www.dangdang.com', 'https://login.dangdang.com/signin.aspx?returnurl=http%3A//www.dangdang.com/'
    ]
banks = [
    'http://www.cmbchina.com', 'https://pbsz.ebank.cmbchina.com/CmbBank_GenShell/UI/GenShellPC/Login/Login.aspx', 'https://pbsz.ebank.cmbchina.com/CmbBank_GenShell/UI/GenShellPC/NetBankAcc/Login.aspx', 
    'http://www.icbc.com.cn', 'https://mybank.icbc.com.cn/icbc/perbank/index.jsp', 'https://corporbank-simp.icbc.com.cn/icbc/normalbank/index.jsp', 'https://vip.icbc.com.cn/icbc/perbank/index.jsp', 
    'http://www.ccb.com', 'https://ibsbjstar.ccb.com.cn/app/V5/CN/STY1/login.jsp', 'https://ibsbjstar.ccb.com.cn/app/V5/CN/STY6/login_pbc.jsp',
    'www.spdb.com.cn', 'https://ebank.spdb.com.cn/per/gb/otplogin.jsp', 'https://ebank.spdb.com.cn/ent/gb/login/osa_query.jsp',   
    'www.abchina.com', 'https://easyabc.95599.cn/SelfBank/netBank/zh_CN/entrance/logonSelf.aspx', 'https://easyabc.95599.cn/commbank/netBank/zh_CN/CommLogin.aspx',
    'https://easyabc.95599.cn/supcorporbank/QryVersionStartUpAct.do', 'https://easyabc.95599.cn/custom/NotCheckStatus/ElecCardLogon.aspx', 
    'http://bank.pingan.com', 'https://www.pingan.com.cn/pinganone/pa/directToMenu.screen?directToMenu=bank_index_index', 'https://ebank.sdb.com.cn/corporbank/logon_basic.jsp',
    'http://www.bankcomm.com', 'https://pbank.95559.com.cn/personbank/logon.jsp', 'https://ebank.95559.com.cn/corporbank/logon.jsp?channel=qywy&oldWay=0', 
    'http://bank.ecitic.com', 'https://e.bank.ecitic.com/perbank5/signIn.do', 'https://enterprise.bank.ecitic.com/corporbank/userLogin.do',
    'http://www.cebbank.com', 'https://www.cebbank.com/per/prePerlogin.do?_locale=zh_CN', 'https://www.cebbank.com/per/preMicrologin.do?_locale=zh_CN',
    'https://ea.cebbank.com/jsp/admin/annuity/Personal_log.jsp', 'https://www.cebbank.com/cebent/prelogin.do?_locale=zh_CN',
    'http://www.cgbchina.com.cn', 'https://ebanks.cgbchina.com.cn/perbank/', 'https://ebank.cgbchina.com.cn/combank/entUserSSLLogin.jsp',
    'http://www.boc.cn', 'https://ebsnew.boc.cn/boc15/login.html', 'https://ebsnew.boc.cn/boc15/login.html?seg=66', 'https://ebs.boc.cn/BocnetClient/LoginFrame3.do?_locale=zh_CN',
    'http://www.psbc.com', 'https://pbank.psbc.com/pweb/prelogin.do?_locale=zh_CN&BankId=9999', 'https://query.psbc.com/eweb/prelogin.do?_locale=zh_CN&BankId=9999&LoginType=Q', 'https://pbank.psbc.com/postremit/remitIndex.do?_locale=zh_CN',
    'http://www.cib.com.cn', 'https://personalbank.cib.com.cn/pers/main/login!changeControlUse.do', 
    'http://www.cmbc.com.cn', 'https://per.cmbc.com.cn/pweb/static/login.html', 'https://per.cmbc.com.cn/pweb/static/blogin.html', 'https://ent.cmbc.com.cn/eweb/static/login.html?fromOldBank=1', 
    'www.bankofbeijing.com.cn', 'https://ebank.bankofbeijing.com.cn/bccbpb/accountLogon.jsp?language=zh_CN', 'https://ebank.bankofbeijing.com.cn/bccbpb/fortuneLogon.jsp',
    'https://ebank.bankofbeijing.com.cn/bccb/corporbank/safelogon.jsp', 'https://ebank.bankofbeijing.com.cn/CBQ/logon.jsp', 'https://ebank.bankofbeijing.com.cn/CBM/logon.jsp',
    'https://ebank.bankofbeijing.com.cn/ebusiness/merchant/logon.jsp', 
    'http://www.hxb.com.cn', 'https://sbank.hxb.com.cn/easybanking/jsp/login/login.jsp', 'https://sbank.hxb.com.cn/gluebanking/login.html?k=true'
    ]
	
	
	
vedio = [
	'http://www.youku.com/','http://v.youku.com/v_show/id_XNzQ1MTI4MTUy.html?qq-pf-to=pcqq.c2c',
	'http://www.tudou.com/','http://www.tudou.com/listplay/wBYvL4DVEO0/fMKYXi1hXdo.html?qq-pf-to=pcqq.c2c',
	'http://www.iqiyi.com/','http://www.iqiyi.com/v_19rrhnm7u4.html',
	'http://tv.sohu.com/','http://tv.sohu.com/yule/','http://tv.sohu.com/20140724/n402655785.shtml',
	'http://v.baidu.com/','http://v.baidu.com/kan/comic/?id=9972&vfm=bdvtx&site=letv.com&n=1&url=http://www.letv.com/ptv/vplay/2223418.html#frp=v.baidu.com/comic_intro/',
	'http://v.qq.com/','http://v.qq.com/voice/','http://v.qq.com/cover/e/e83qn96pujrypsz.html?vid=g0132w8xj3l'
	'http://www.letv.com/','http://www.letv.com/ptv/vplay/20234078.html?ref=ym0202'
	
	'http://www.1ting.com/p_216341_118402_216146_118175_173633_216137_165126_1044416_99706_118404_226850_1044450_1044430_1044412_1044426_1044439_216144_226849_1044414_1044417_462747_1044415_462751_1044413_89271_1044418_462746_173635_216140_1044433.html',
	'http://baidu.hz.letv.com/kan/TQnr?fr=v.baidu.com/',
	'http://www.1ting.com/p_89271.html',
	'http://www.1ting.com/',
	'http://www.1ting.com/player/65/player_118175.html',
	'http://www.xinhuanet.com/newscenter/index.htm',
	'http://www.letv.com/ptv/vplay/20253246.html'
	]
	
	
#websites + banks
websites = websites + banks + vedio
WebsitesTest(websites)


wordFileToOpen = os.getcwdu() + '\\test.doc'
wordFileToSave = os.getcwdu() + '\\savedfiles\\test_doc'
WinwordTest(wordFileToOpen, wordFileToSave)

excelFileToOpen = os.getcwdu() + '\\test.xls'
excelFileToSave = os.getcwdu() + '\\savedfiles\\test_xls'
ExcelTest(excelFileToOpen, excelFileToSave)


pptFileToOpen = os.getcwdu() + '\\test.ppt'
pptFileToSave = os.getcwdu() + '\\savedfiles\\test_ppt'
PowerpointTest(pptFileToOpen, pptFileToSave)




pdfFileToOpen = os.getcwdu() + '\\test.pdf'
AcrobatRdTest(pdfFileToOpen, 'test.pdf')


wpsFileToOpen = os.getcwdu() + '\\test.doc'
wpsFileToSave = os.getcwdu() + '\\savedfiles\\test_wps.txt'
WpsTest(wpsFileToOpen, wpsFileToSave)

mp3FileToOpen = os.getcwdu() + '\\KissFromARose.mp3'
WMPlayerTest(mp3FileToOpen)

Rundll32Test()

mailFileToSave = os.getcwdu() + '\\savedfiles\\test_mail'
OutlookTest(mailFileToSave)










