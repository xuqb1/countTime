#COMPILER PBWIN 9
#COMPILE EXE
#DIM ALL

%USEMACROS = 1

#RESOURCE "countTime.PBR"
#INCLUDE "WIN32API.INC"
#INCLUDE "COMMCTRL.INC"
%ID_TIMER1        = 100
%IDC_COUNTSTATE   = 120    '计时状态：准备，1/3/5分钟计时中...，
%IDC_TIMESHOW     = 121      '倒数牌
%IDC_ONEMINUTE    = 122     '1分钟
%IDC_TWOMINUTE    = 123     '2分钟
%IDC_THREEMINUTE  = 124   '3分钟
%IDC_FIVEMINUTE   = 125    '5分钟
%IDC_RESET        = 126         '重置
%IDC_CUSMINUTE    = 101     '自定义分钟数
%IDC_STARTBUT     = 102
%IDC_CUSMINUTELB  = 103     '分钟标签
%IDC_MSCTLS_UPDOWN32=104
%IDC_SETBT        = 105
%IDC_WORKTIMELB   = 106
%IDC_WORKTIMETB   = 107
%IDC_RESTTIMELB   = 108
%IDC_RESTTIMETB   = 109
%IDC_WORKTIPLB   = 110
%IDC_WORKTIPTB   = 111
%IDC_RESTTIPLB   = 112
%IDC_RESTTIPTB   = 113
%IDC_MINUTECB    = 114
%IDC_NOWTIMELB   = 115 '当前时间
%ID_TIMER2      = 116
%WM_TRAY = 105
$SOFTNAME = " 久坐/休息计时器"
$SOFTVERSION ="v1"
GLOBAL hMenu AS LONG
GLOBAL firsttime AS LONG
GLOBAL secondtime AS LONG
GLOBAL firststr AS STRING
GLOBAL secondstr AS STRING
GLOBAL gAppName AS STRING
GLOBAL gAppVersion AS STRING
GLOBAL minute() AS STRING
GLOBAL bgColor AS LONG
'====================================================================
FUNCTION PBMAIN () AS LONG
  LOCAL hDlg AS DWORD
  LOCAL hFont1 AS DWORD
  LOCAL xStr AS STRING
  LOCAL yStr AS STRING
  LOCAL deskWidth AS LONG
  LOCAL deskHeight AS LONG
  LOCAL dlgWidth AS LONG
  LOCAL dlgHeight AS LONG
  'local bgColor as long

  'bgColor=RGB(0,120,215)
  bgColor=RGB(0,66,117)

  dlgWidth=203
  dlgHeight=110
  gAppName=GetConfig(EXE.PATH$ & "config.ini","default","appname",$SOFTNAME)
  gAppVersion = GetConfig(EXE.PATH$ & "config.ini","default","version",$SOFTVERSION)
  xStr=GetConfig(EXE.PATH$ & "config.ini","default","xpos","")
  yStr=GetConfig(EXE.PATH$ & "config.ini","default","ypos","")
  DESKTOP GET CLIENT TO deskWidth,deskHeight
  firsttime=VAL(GetConfig(EXE.PATH$ & "config.ini","default","worktime","40"))
  secondtime=VAL(GetConfig(EXE.PATH$ & "config.ini","default","resttime","10"))
  REDIM minute(1)
  minute(0)="work,"+FORMAT$(firsttime)
  minute(1)="rest,"+FORMAT$(secondtime)
  firststr=GetConfig(EXE.PATH$ & "config.ini","default","worktip","工作中")
  secondstr=GetConfig(EXE.PATH$ & "config.ini","default","resttip","休息中")
  DIALOG NEW  0,  $SOFTNAME & " " & $SOFTVERSION,,, dlgWidth, dlgHeight,  %WS_CAPTION OR %WS_SYSMENU OR %WS_MINIMIZEBOX, %WS_EX_TOPMOST TO hDlg
  DIALOG SET ICON hDlg,"APPICON"
  DIALOG PIXELS hDlg,deskWidth,deskHeight TO UNITS deskWidth,deskHeight
  IF xStr="" OR VAL(xStr)<0 THEN
    xStr= STR$((deskWidth-dlgWidth)\2)
  END IF
  IF VAL(xStr)+dlgWidth>deskWidth THEN
    xStr=STR$(deskWidth-dlgWidth)
  END IF
  IF yStr="" OR VAL(yStr)<0 THEN
    yStr=STR$((deskHeight-dlgHeight)\2)
  END IF
  IF VAL(yStr)+dlgHeight>deskHeight THEN
    yStr=STR$(deskHeight-dlgHeight)
  END IF
  DIALOG SET LOC hDlg,VAL(xStr),VAL(yStr)
  SetConfig(EXE.PATH$ & "config.ini","default","xpos",xStr)
  SetConfig(EXE.PATH$ & "config.ini","default","ypos",yStr)
  DIALOG SET COLOR hDlg, %WHITE, bgColor 'rgb(0,150,215) '%BLACK
  CONTROL ADD LABEL ,hDlg,%IDC_COUNTSTATE, "准备",5,5,150,12,%SS_SIMPLE  ,%WS_EX_TRANSPARENT
  CONTROL SET COLOR hDlg,%IDC_COUNTSTATE,%WHITE,bgColor '%BLACK
  FONT NEW "font1",12, 1, 136, 0, 0 TO hFont1
  CONTROL SET FONT hDlg,%IDC_COUNTSTATE,hFont1
  FONT END hFont1

  CONTROL ADD LABEL ,hDlg,%IDC_NOWTIMELB, "11:32:20",20,55,120,30,%SS_SIMPLE  ,%WS_EX_TRANSPARENT
  CONTROL SET COLOR hDlg,%IDC_NOWTIMELB,%WHITE,bgColor '%BLACK
  FONT NEW "font1",30, 1, 1, 0, 0 TO hFont1
  CONTROL SET FONT hDlg,%IDC_NOWTIMELB,hFont1
  FONT END hFont1

  CONTROL ADD LABEL ,hDlg,%IDC_TIMESHOW,"00", 30,15,130,40
  CONTROL SET COLOR hDlg,%IDC_TIMESHOW,%WHITE,bgColor '%BLACK
  FONT NEW "font1",50, 1, 136, 0, 0 TO hFont1
  CONTROL SET FONT hDlg,%IDC_TIMESHOW,hFont1
  FONT END hFont1
  CONTROL ADD COMBOBOX,hDlg,%IDC_MINUTECB,minute(),100,90,50,60
  CONTROL SET COLOR hDlg,%IDC_MINUTECB,%WHITE,bgColor '%BLACK
  CONTROL SET TEXT hDlg,%IDC_MINUTECB,minute(0)
  'CONTROL ADD BUTTON ,hDlg,%IDC_ONEMINUTE,"1分钟",5,75,30,20
  'CONTROL ADD BUTTON ,hDlg,%IDC_TWOMINUTE,"2分钟",40,75,30,20
  'CONTROL ADD BUTTON ,hDlg,%IDC_THREEMINUTE,"3分钟",75,75,30,20
  'CONTROL ADD BUTTON ,hDlg,%IDC_FIVEMINUTE,"5分钟",110,75,30,20
  CONTROL ADD BUTTON ,hDlg,%IDC_RESET,"停止",165,50,30,20
  'CONTROL ADD TEXTBOX, hDlg, %IDC_CUSMINUTE, "30", 110,58,20,12, %WS_CHILD _
  '                      OR %WS_VISIBLE, %WS_EX_CLIENTEDGE
  'CONTROL ADD "msctls_updown32", hDlg, %IDC_MSCTLS_UPDOWN32, _
  '                      "msctls_updown32_4", 130, 58, 12, 12,  %WS_CHILD OR %WS_VISIBLE OR _
  '                      %UDS_SETBUDDYINT OR %UDS_AUTOBUDDY OR %UDS_WRAP
  'CONTROL SEND hDlg, %IDC_MSCTLS_UPDOWN32, %UDM_SETRANGE32, 1  ,600
  'CONTROL SEND hDlg, %IDC_MSCTLS_UPDOWN32, %UDM_SETPOS32  ,0,30
  'CONTROL ADD LABEL  ,hDlg,%IDC_CUSMINUTELB,"分钟",143,59,25,10
  CONTROL SET COLOR hDlg,%IDC_CUSMINUTELB,%WHITE,bgColor '%BLACK
  FONT NEW "font1",12, 1, 136, 0, 0 TO hFont1
  CONTROL SET FONT hDlg,%IDC_CUSMINUTELB,hFont1
  FONT END hFont1
  CONTROL ADD BUTTON ,hDlg,%IDC_STARTBUT,"开始",165,25,30,20
  CONTROL ADD BUTTON ,hDlg,%IDC_SETBT,"设置",165,75,30,20
  MENU NEW POPUP TO hMenu
  MENU ADD STRING, hMenu, "主窗口", 401, %MF_ENABLED
'  MENU ADD STRING, hMenu, "-", 491, %MF_SEPARATOR
'  MENU ADD STRING, hMenu, "Cancel", 491, %MF_ENABLED
'  MENU ADD STRING, hMenu, "-", 491, %MF_SEPARATOR
  MENU ADD STRING, hMenu, "关于", 402, %MF_ENABLED
  MENU ADD STRING, hMenu, "退出", 403, %MF_ENABLED
  DIALOG SHOW MODAL hDlg, CALL DlgProc
END FUNCTION
'====================================================================
CALLBACK FUNCTION DlgProc() AS LONG
  LOCAL ac, c AS LONG, sTime AS STRING
  LOCAL tmpStr AS STRING
  STATIC ti AS NOTIFYICONDATA
  STATIC p AS POINTAPI
  STATIC hIcon1 AS LONG
  STATIC hIcon2 AS LONG
  STATIC iconFlag AS LONG
  LOCAL xpos,ypos AS LONG
  SELECT CASE CB.MSG
    CASE %WM_INITDIALOG                             ' <- sent right before the dialog is displayed.
      STATIC countType AS INTEGER '计时类型:1 ,3,5
      STATIC maxSec AS INTEGER '最大秒数
      STATIC minSec AS INTEGER '最小秒数
      STATIC curSec AS INTEGER '当前秒数
      STATIC idEvent AS LONG
      STATIC idEvent2 AS LONG
      iconFlag=0
      hIcon1=LoadImage(0, "time.ico", %IMAGE_ICON, 0, 0, %LR_LOADFROMFILE)
      hIcon2=LoadImage(0, "time1.ico", %IMAGE_ICON, 0, 0, %LR_LOADFROMFILE)
      ti.cbSize = SIZEOF(ti)
      ti.hWnd = CB.HNDL
      ti.uID = GetWindowLong(CB.HNDL,%GWL_HINSTANCE)   'hInst
      ti.uFlags = %NIF_ICON OR %NIF_MESSAGE OR %NIF_TIP
      ti.uCallbackMessage = %WM_TRAY
      ti.hIcon = hIcon1'LoadIcon(%NULL,"APPICON")
      ti.szTip = $SOFTNAME & $SOFTVERSION
      Shell_NotifyIcon %NIM_ADD, ti

      idEvent2 = SetTimer(CB.HNDL, %ID_TIMER2, _    ' Create WM_TIMER events with the SetTimer API
                           1000, BYVAL %NULL)
      DIALOG POST CB.HNDL, %WM_TIMER, %ID_TIMER2, 0
    CASE %WM_TRAY
      SELECT CASE AS LONG LOWRD(CB.LPARAM)
        CASE %WM_RBUTTONDOWN
          SetForegroundWindow CB.HNDL
          GetCursorPos p
          TrackPopupMenu hMenu, %TPM_BOTTOMALIGN OR %TPM_RIGHTALIGN,p.x,p.y, 0, CB.HNDL, BYVAL %NULL
        ' Popup main dialog (double left click on tray)
        CASE %WM_LBUTTONDBLCLK
          ' Another way to show the dialog, cancel the flashing icon
          ' and restore the default tray icon
          'KillTimer cb.hndl, %TIMER_1
          ti.hIcon = hIcon1 'LoadIcon(%NULL,BYVAL MAKLNG(%IDI_QUESTION,0))
          Shell_NotifyIcon %NIM_MODIFY, ti

          'CONTROL SET TEXT hDlg&, 110, UDPMessage
          DIALOG SHOW STATE CB.HNDL,%SW_RESTORE  'hDlg& CALL CB_Dlg
          CONTROL SET FOCUS CB.HNDL,%IDC_TIMESHOW

      END SELECT
    CASE %WM_TIMER                                  ' Posted by the created timer
      SELECT CASE CB.WPARAM
        CASE %ID_TIMER1              ' Make sure it's corrent timer id
          IF curSec<=minSec+10 THEN
            BEEP
            BEEP
          END IF
          IF iconFlag=0 THEN
            ti.hIcon=hIcon2
            iconFlag=1
          ELSE
            iconFlag=0
            ti.hIcon=hIcon1
          END IF
          ti.szTip = $SOFTNAME & $SOFTVERSION  & STR$(curSec)
          Shell_NotifyIcon %NIM_MODIFY, ti
          IF curSec<minSec THEN
            KillTimer CB.HNDL, idEvent
            DIALOG SHOW STATE CB.HNDL,%SW_RESTORE
            CONTROL SET FOCUS CB.HNDL,%IDC_TIMESHOW
            'CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE,'FORMAT$(countType) & " 分钟计时结束"
            'EnableCtl CB.HNDL
            IF maxSec=firsttime*60 THEN
              countType=secondtime
              maxSec=secondtime*60
              CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE, secondstr
              CONTROL SET TEXT CB.HNDL,%IDC_MINUTECB,"rest," & FORMAT$(secondtime)
            ELSE
              countType=firsttime
              maxSec=firsttime*60
              CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE, firststr
              CONTROL SET TEXT CB.HNDL,%IDC_MINUTECB,"work," & FORMAT$(firsttime)
            END IF
            curSec=maxSec
            idEvent = SetTimer(CB.HNDL, %ID_TIMER1, _    ' Create WM_TIMER events with the SetTimer API
                             1000, BYVAL %NULL)
            DIALOG POST CB.HNDL, %WM_TIMER, %ID_TIMER1, 0
            EXIT FUNCTION
          END IF
          IF curSec<5 AND curSec>0 THEN
            CONTROL SET COLOR CB.HNDL,%IDC_TIMESHOW,%YELLOW,bgColor '%BLACK
          ELSEIF curSec<=0 THEN
            CONTROL SET COLOR CB.HNDL,%IDC_TIMESHOW,%RED,bgColor '%BLACK
          ELSE
            CONTROL SET COLOR CB.HNDL,%IDC_TIMESHOW,%WHITE,bgColor '%BLACK
          END IF
          CONTROL SET TEXT CB.HNDL,%IDC_TIMESHOW,FORMAT$(curSec)
          DECR curSec
        CASE %ID_TIMER2
          SetNowTime CB.HNDL
      END SELECT
    CASE %WM_DESTROY                                ' Sent when the dialog is being destroyed
      IF idEvent THEN                             ' If a timer identifier exists
        KillTimer CB.HNDL, idEvent               ' make sure to stop the timer events
      END IF
      IF idEVent2 THEN
        KillTimer CB.HNDL, idEVent2
      END IF
      Shell_NotifyIcon %NIM_DELETE, ti
      DIALOG GET LOC CB.HNDL TO xpos,ypos
      SetConfig(EXE.PATH$ & "config.ini","default","xpos",STR$(xpos))
      SetConfig(EXE.PATH$ & "config.ini","default","ypos",STR$(ypos))
      SetConfig(EXE.PATH$ & "config.ini","default","appname",$SOFTNAME)
      SetConfig(EXE.PATH$ & "config.ini","default","version",$SOFTVERSION)
    CASE %WM_COMMAND       ' <- A control is calling
      IF CB.CTLMSG <> %BN_CLICKED THEN
        EXIT FUNCTION
      END IF
      SELECT CASE CB.CTL  ' <- Look at control's id
        CASE 401 '主窗口菜单
          ' First, update the dialog message with latest received message
          'CONTROL SET TEXT hDlg&, 110, UDPMessage
          DIALOG SHOW STATE CB.HNDL,%SW_RESTORE 'hDlg& CALL CB_Dlg
          CONTROL SET FOCUS CB.HNDL,%IDC_TIMESHOW
        ' About
        CASE 402 '关于菜单
            MSGBOX $SOFTNAME & $SOFTVERSION & " by xsoft" _
            & CHR$(13,10,10) & "本程序可自由复制使用。" _
            & CHR$(13,10) & "Enjoy!", %MB_OK + %MB_ICONINFORMATION,"关于"
        ' Quit
        CASE 403 '退出菜单
            DIALOG END CB.HNDL
        CASE %IDC_STARTBUT '开始按钮
          'CONTROL GET TEXT CB.HNDL,%IDC_CUSMINUTE TO tmpStr
          'IF tmpStr="" THEN
          '  tmpStr="2"
          'END IF
          CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE,firststr 'tmpStr & " 分钟计时中..."
          DisableCtl CB.HNDL
          CONTROL GET TEXT CB.HNDL,%IDC_MINUTECB TO tmpStr
          IF PARSE$(tmpStr,",",1)="work" THEN
            firsttime=VAL(PARSE$(tmpStr,",",2))
            countType = firsttime
            maxSec = firsttime*60
          ELSE
            secondtime=VAL(PARSE$(tmpStr,",",2))
            countType = secondtime
            maxSec = secondtime*60
            CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE,secondstr
          END IF

          minSec = 0
          curSec=maxSec
          idEvent = SetTimer(CB.HNDL, %ID_TIMER1, _    ' Create WM_TIMER events with the SetTimer API
                           1000, BYVAL %NULL)
          DIALOG POST CB.HNDL, %WM_TIMER, %ID_TIMER1, 0
        CASE %IDC_RESET
          KillTimer CB.HNDL, idEvent
          CONTROL SET TEXT CB.HNDL,%IDC_COUNTSTATE,"准备"
          CONTROL SET TEXT CB.HNDL,%IDC_TIMESHOW,"00"
          EnableCtl CB.HNDL
        CASE %IDC_SETBT
          ShowSettingDlg CB.HNDL
      END SELECT
    CASE %WM_SYSCOMMAND
      SELECT CASE (CB.WPARAM AND &H0FFF0)
        CASE %SC_MINIMIZE
          ShowWindow CB.HNDL, %SW_HIDE
          FUNCTION = 1
          EXIT FUNCTION
        CASE %SC_CLOSE
          Shell_NotifyIcon %NIM_DELETE, ti
      END SELECT
  END SELECT
END FUNCTION
FUNCTION SetNowTime(BYVAL hWnd AS DWORD)AS LONG
  CONTROL SET TEXT hWnd,%IDC_NOWTIMELB, TIME$
END FUNCTION
SUB DisableCtl(BYVAL hWnd AS DWORD)
  CONTROL DISABLE hWnd, %IDC_CUSMINUTE
  CONTROL DISABLE hWnd, %IDC_MSCTLS_UPDOWN32
  CONTROL DISABLE hWnd, %IDC_STARTBUT
  CONTROL DISABLE hWnd, %IDC_ONEMINUTE
  CONTROL DISABLE hWnd, %IDC_TWOMINUTE
  CONTROL DISABLE hWnd, %IDC_THREEMINUTE
  CONTROL DISABLE hWnd, %IDC_FIVEMINUTE
  CONTROL DISABLE hWnd, %IDC_MINUTECB
  DIALOG REDRAW hWnd
END SUB
SUB EnableCtl(BYVAL hWnd AS DWORD)
  CONTROL ENABLE hWnd, %IDC_ONEMINUTE
  CONTROL ENABLE hWnd, %IDC_TWOMINUTE
  CONTROL ENABLE hWnd, %IDC_THREEMINUTE
  CONTROL ENABLE hWnd, %IDC_FIVEMINUTE
  CONTROL ENABLE hWnd, %IDC_CUSMINUTE
  CONTROL ENABLE hWnd, %IDC_MSCTLS_UPDOWN32
  CONTROL ENABLE hWnd, %IDC_STARTBUT
  CONTROL ENABLE hWnd, %IDC_MINUTECB
  DIALOG REDRAW hWnd
END SUB
FUNCTION ShowSettingDlg(BYVAL hParent AS DWORD)AS LONG
  LOCAL hDlg AS DWORD
  DIALOG NEW  hParent,  "设置",,, 160, 65,  %WS_CAPTION OR %WS_SYSMENU TO hDlg
  CONTROL ADD LABEL,hDlg,%IDC_WORKTIMELB,"工作时长(分钟):",5,6,55,10
  CONTROL ADD TEXTBOX,hDlg,%IDC_WORKTIMETB,STR$(firsttime),62,5,15,12,%ES_NUMBER OR %WS_BORDER OR %WS_TABSTOP
  CONTROL ADD LABEL,hDlg,%IDC_WORKTIPLB,"提示:",80,6,18,10
  CONTROL ADD TEXTBOX,hDlg,%IDC_WORKTIPTB,firststr,100,5,50,12

  CONTROL ADD LABEL,hDlg,%IDC_RESTTIMELB,"休息时长(分钟):",5,20,55,10
  CONTROL ADD TEXTBOX,hDlg,%IDC_RESTTIMETB,STR$(secondtime),62,19,15,12,%ES_NUMBER OR %WS_BORDER OR %WS_TABSTOP
  CONTROL ADD LABEL,hDlg,%IDC_RESTTIPLB,"提示:",80,20,18,10
  CONTROL ADD TEXTBOX,hDlg,%IDC_RESTTIPTB,secondstr,100,19,50,12

  CONTROL ADD BUTTON,hDlg,%IDOK,"确定",38,40,40,15
  CONTROL ADD BUTTON,hDlg,%IDCANCEL,"关闭",83,40,40,15
  DIALOG SHOW MODAL hDlg, CALL SettingDlgProc
END FUNCTION
CALLBACK FUNCTION SettingDlgProc
  LOCAL tmpstr AS STRING
  SELECT CASE CB.MSG
    CASE %WM_INITDIALOG

    CASE %WM_COMMAND
      IF CB.CTLMSG <> %BN_CLICKED THEN
        EXIT FUNCTION
      END IF
      SELECT CASE CB.CTL  ' <- Look at control's id
        CASE %IDOK
          CONTROL GET TEXT CB.HNDL,%IDC_WORKTIMETB TO tmpstr
          IF TRIM$(tmpstr)="" THEN
            MSGBOX "请输入工作时长",%MB_ICONINFORMATION,"提示"
            CONTROL SET FOCUS CB.HNDL,%IDC_WORKTIMETB
            EXIT FUNCTION
          END IF
          firsttime=VAL(TRIM$(tmpstr))
          SetConfig EXE.PATH$ & "config.ini","default","worktime",STR$(firsttime)
          CONTROL GET TEXT CB.HNDL,%IDC_WORKTIPTB TO tmpstr
          firststr=TRIM$(tmpstr)
          SetConfig EXE.PATH$ & "config.ini","default","worktip",firststr

          CONTROL GET TEXT CB.HNDL,%IDC_RESTTIMETB TO tmpstr
          IF TRIM$(tmpstr)="" THEN
            MSGBOX "请输入休息时长",%MB_ICONINFORMATION,"提示"
            CONTROL SET FOCUS CB.HNDL,%IDC_RESTTIMETB
            EXIT FUNCTION
          END IF
          secondtime=VAL(TRIM$(tmpstr))
          SetConfig EXE.PATH$ & "config.ini","default","resttime",STR$(secondtime)
          CONTROL GET TEXT CB.HNDL,%IDC_RESTTIPTB TO tmpstr
          secondstr=TRIM$(tmpstr)
          SetConfig EXE.PATH$ & "config.ini","default","resttip",secondstr
          DIALOG END CB.HNDL,0
        CASE %IDCANCEL
          DIALOG END CB.HNDL,0
      END SELECT
  END SELECT
END FUNCTION
'-----------------------------
' 获取配置
'-----------------------------
FUNCTION getConfig(BYVAL FileName AS STRING,BYVAL SecName AS STRING,BYVAL keynameStr AS STRING,BYVAL defaultVal AS STRING)AS STRING
  LOCAL IniFile       AS ASCIIZ * 256
  LOCAL SectionName   AS ASCIIZ * 256
  LOCAL KeyName       AS ASCIIZ * 256
  LOCAL defaultValue  AS ASCIIZ * 256
  LOCAL resultStr     AS ASCIIZ * 256
  IniFile = FileName
  SectionName = SecName
  KeyName = keynameStr
  defaultValue = defaultVal
  GetPrivateProfileString SectionName,KeyName,defaultValue,resultStr,256,IniFile
  'msgbox KeyName & "=" & resultStr
  FUNCTION=resultStr
END FUNCTION
'---------------------------------
' 设置配置
'---------------------------------
FUNCTION setConfig(BYVAL FileName AS STRING,BYVAL SecName AS STRING,BYVAL keyname AS STRING,BYVAL sValue AS STRING)AS LONG
  LOCAL SecNameAsc AS ASCIIZ * 256
  LOCAL keynameAsc AS ASCIIZ * 256
  LOCAL sValueAsc AS ASCIIZ * 256
  LOCAL FileNameAsc AS ASCIIZ * 256
  SecNameAsc=SecName
  keynameAsc=keyname
  sValueAsc=sValue
  FileNameAsc=FileName
  WritePrivateProfileString SecNameAsc,keynameAsc,sValueAsc,FileNameAsc
END FUNCTION
