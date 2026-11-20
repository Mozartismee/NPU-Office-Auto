@echo off
echo ************************************************************************
echo ************************************************************************
echo **       本程式將執行 office 2013 2016 2019 2021 2024                 **
echo **                    等多種版本的啟動認證的設定                      **
echo **       請確認您執行本程式時是否已經選擇已關閉  [防毒軟體]           **
echo **       請確認您執行本程式時是否已經選擇以  [系統管理者身分執行]     **
echo **       如果沒有的話請點選本程式圖示在按滑鼠右鍵即可選擇執行方式     **
echo ************************************************************************
echo ************************************************************************
echo          鍵盤(1)OFFICE 2013(校內環境)  (6)OFFICE 2013 (校外VPN)
echo          鍵盤(2)OFFICE 2016(校內環境)  (7)OFFICE 2016 (校外VPN)
echo          鍵盤(3)OFFICE 2019(校內環境)  (8)OFFICE 2019 (校外VPN)
echo          鍵盤(4)OFFICE 2021(校內環境)  (9)OFFICE 2021 (校外VPN)
echo          鍵盤(5)OFFICE 2024(校內環境)  (0)OFFICE 2024 (校外VPN)
set /p input=  請選擇 1 , 2 , 3 , 4 , 5 , 6 , 7 , 8 .....?
if (%input%) == () goto start
set input=%input:~0,1%
if (%input%) == (0) goto add9
if (%input%) == (5) goto add8
if (%input%) == (9) goto add7
if (%input%) == (8) goto add6
if (%input%) == (7) goto add5
if (%input%) == (6) goto add4
if (%input%) == (4) goto add3
if (%input%) == (3) goto add2
if (%input%) == (2) goto add1
if (%input%) == (1) goto add
:add
@echo off
echo 歡迎使用Office KMS 金鑰啟動程式。
echo 偵測 Microsoft Office 2013 安裝目錄。
set ospp="%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs"
if exist %ospp% goto ospp_act
set ospp="%ProgramFiles%\Microsoft Office\Office15\ospp.vbs"
if exist %ospp% goto ospp_act
echo Office 啟動失敗，找不到Office 2013安裝位置
pause
goto end


:ospp_act
echo 正在指定Office KMS 金鑰管理伺服器...
cscript %ospp% /sethst:210.70.253.10
echo 啟動 Microsoft Office 2013
cscript %ospp% /setprt:1688
cscript %ospp% /act


echo 啟動程序執行完成
echo 上方需有 Product activation successful 文字,
echo 才表示您的 Office2013 啟動成功!

pause

goto end
:add1
@echo off
echo 偵測 Microsoft Office 2016 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:210.70.253.10
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2016
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2016 啟動成功!
pause

goto end
:add2
@echo off
echo 偵測 Microsoft Office 2019 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:210.70.253.10
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2019
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2019 啟動成功!
pause

goto end
:add3
@echo off
echo 偵測 Microsoft Office 2021 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:210.70.253.10
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2021
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2021 啟動成功!
pause
goto end

:add4
@echo off
echo 歡迎使用Office KMS 金鑰啟動程式。
echo 偵測 Microsoft Office 2013 安裝目錄。
set ospp="%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs"
if exist %ospp% goto ospp_act
set ospp="%ProgramFiles%\Microsoft Office\Office15\ospp.vbs"
if exist %ospp% goto ospp_act
echo Office 啟動失敗，找不到Office 2013安裝位置
pause
goto end


:ospp_act
echo 正在指定Office KMS 金鑰管理伺服器...
cscript %ospp% /sethst:192.168.168.1
echo 啟動 Microsoft Office 2013
cscript %ospp% /setprt:1688
cscript %ospp% /act


echo 啟動程序執行完成
echo 上方需有 Product activation successful 文字,
echo 才表示您的 Office2013 啟動成功!

pause

goto end
:add5
@echo off
echo 偵測 Microsoft Office 2016 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:192.168.168.1
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2016
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2016 啟動成功!
pause

goto end
:add6
@echo off
echo 偵測 Microsoft Office 2019 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:192.168.168.1
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2019
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2019 啟動成功!
pause

goto end
:add7
@echo off
echo 偵測 Microsoft Office 2021 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:192.168.168.1
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2021
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2021 啟動成功!
pause

goto end
:add8
@echo off
echo 偵測 Microsoft Office 2024 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:210.70.253.10
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2024
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2024 啟動成功!
pause


goto end
:add9
@echo off
echo 偵測 Microsoft Office 2024 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo 重啟 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo 設定 KMS 金鑰管理伺服器
cscript %OfficePath%ospp.vbs /sethst:192.168.168.1
cscript %OfficePath%ospp.vbs /setprt:1688
echo 啟動 Microsoft Office 2024
cscript %OfficePath%ospp.vbs /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現,
echo 才表示您的 Office2024 啟動成功!
pause


goto end
:end

