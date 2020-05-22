@echo off
mode con cols=100 lines=30
title E龙のKMS激活脚本V3.1（果核剥壳KMS）
setlocal EnableDelayedExpansion&color F8 & cd /d "%~dp0"
%1 %2
ver|find "5.">nul&&goto :UAC
mshta vbscript:createobject("shell.application").shellexecute("%~s0","goto :UAC","","runas",1)(window.close)&goto :eof

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
:UAC
if '%errorlevel%' NEQ '0' (
    echo 请求管理员权限...
    goto UACPrompt
) else ( goto gotAdmin )
:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\ghpymkmsgetadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\ghpymkmsgetadmin.vbs"
    "%temp%\ghpymkmsgetadmin.vbs"
    exit /B
:gotAdmin
    if exist "%temp%\ghpymkmsgetadmin.vbs" ( del "%temp%\ghpymkmsgetadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
    goto start

:start
cls
echo E龙のKMS激活脚本V3.1（果核剥壳KMS）
echo.&echo 果核剥壳KMS地址 kms.ghpym.com/kms.qkeke.com
echo.&echo 正在检查与激活服务器的连接情况……
set KMSSERVER=kms.ghpym.com
echo.
ping -n 1 %KMSSERVER% | find "超时" >nul && goto network_error
ping -n 1 %KMSSERVER% | find "目标主机" >nul && goto network_error
ping -n 1 %KMSSERVER% | find "无法访问" >nul && goto network_error
ping -n 1 %KMSSERVER% | find "故障" >nul && goto network_error
ping -n 1 %KMSSERVER% | find "找不到" >nul && goto network_error
echo 成功连接上服务器！按任意键以继续…… & pause>nul
goto begin

:network_error
echo 无法连接到服务器QwQ，按任意键强制继续…… & pause>nul
goto begin

:begin
cls
echo —————注意—————
echo 本脚本需要管理员权限运行
echo —————注意—————
echo 仅支持批量许可的Office/Windows才能激活,180天有效期，联网7天自动循环续期
echo 留言/转载请访问edragon.sxl.cn或发邮件至e_dragon007@qq.com，转载请留下转载链接,我会将链接加入程序
echo 本脚本使用了 @果核剥壳 的KMS激活服务器，链接：ghpym.com
echo 本程序发布在github且遵循GPL-3.0协议免费发行，如需修改需声明修改处，注明原作者并以相同许可证发行！
echo 　　　　　　　　　　　　██████　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
echo 　　　　　██　　　　██　█████　　　　　　　　　　　　　█　　　　　　　　　　　　　　　　
echo 　　　　████　　　████████　　　　　　　　　██　███　　　　　　　　　　　　　　　
echo 　　　　　　　　　　　████████　　　　　　　　████　　　　　　　　　　　　　　　　　　
echo 　　　　　　　　　　　████　　　　　　　　　　　　　　　　　　　　　　　　█　　　　　　　　　
echo 　　　　　　　　　　　███████　　　　　　　　　　　　　　　　　　　　　█　　　　　　　　　
echo 　　　　█　　　　　█████　　　　　　　　　　　　　　　　　　　　　　█　█　█　　　　　　　
echo 　　　　█　　　　██████　　　　　　　　　　　　　　　　　　　　　　█　█　█　　　　　　　
echo 　　　　██　　　████████　　　　　　　　　　　　　　　　　　　　█　█　█　　　　　　　
echo 　　　　███　███████　█　　　　　　　　　　　　　　　　　　　　█　█　█　　　　　　　
echo 　　　　███████████　　　　　　　　　　　　　　　　　　　　　　　███　　　　　　　　
echo 　　　　　█████████　　　　　　　　　　　　　　　　　　　　　　　　　█　　　　　　　　　
echo 　　　　　　████　██　　　　　　　　　　　　　　　　▂▄　　　　　　　　█　　　　　　　　　
echo █████　　　█　　█　　███████████████████████████████████
echo 　　　　　　　　██　██　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
echo [1]使用果核剥壳KMS激活Windows                     [6]清除Office KMS激活
echo [2]使用果核剥壳KMS激活Office                      [7]查询Windows激活状态
echo [3]使用其他KMS服务器激活Windows                   [8]查询Office激活状态
echo [4]使用其他KMS服务器激活Office                    [9]访问我的个人主页
echo [5]清除Windows KMS激活                            [0]查看软件发布帖
set /p choose=请输入你的选择:
if /i "%choose%"=="1" cls&goto choose1
if /i "%choose%"=="2" cls&goto choose2
if /i "%choose%"=="3" cls&goto choose3
if /i "%choose%"=="4" cls&goto choose4
if /i "%choose%"=="5" cls&goto choose5
if /i "%choose%"=="6" cls&goto choose6
if /i "%choose%"=="7" cls&goto choose7
if /i "%choose%"=="8" cls&goto choose8
if /i "%choose%"=="9" cls&goto choose9
if /i "%choose%"=="0" cls&goto choose0
exit

:choose1
set KMSSERVER=kms.ghpym.com
goto choose1a

:choose2
set KMSSERVER=kms.ghpym.com
goto choose2a

:choose1a
ver | find "6.0." > nul && goto winvista
ver | find "6.1." > nul && goto win7
ver | find "6.2." > nul && goto win8
ver | find "6.3." > nul && goto win81
ver | find "10.0." > nul && goto win10
echo 未找到合适的系统QwQ，按任意键返回…… & pause>nul
goto begin

:winvista
echo 当前系统为Windows Vista/2008。
set Business=YFKBB-PQJJV-G996G-VWGXY-2V3X8
set BusinessN=HMBQG-8H2RH-C77VX-27R82-VMQBT
set Enterprise=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
set EnterpriseN=VTC42-BM838-43QHV-84HX6-XJXKV
set ServerWeb=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
set ServerStandard=TM24T-X9RMF-VWXK6-X8JC9-BFGM2
set ServerStandardV=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
set ServerEnterprise=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
set ServerEnterpriseV=39BXF-X8Q23-P2WWT-38T2F-G3FPG
set ServerWeb=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
set ServerDatacenter=7M67G-PC374-GR742-YH8V4-TCBY3
set ServerDatacenterV=22XQ2-VRXRG-P8D42-K34TD-G3QQC
set ServerEnterpriseIA64=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto winact

:win7
echo 当前系统为Windows 7/2008 R2。
set Professional=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
set ProfessionalN=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
set ProfessionalE=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
set Enterprise=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
set EnterpriseN=YDRBP-3D83W-TY26F-D46B2-XCKRJ
set EnterpriseE=C29WB-22CC8-VJ326-GHFJW-H9DH4
set ServerWeb=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
set ServerHPC=TT8MH-CG224-D3D7Q-498W2-9QCTX
set ServerStandard=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
set ServerEnterprise=489J6-VHDMP-X63PK-3K798-CPX3Y
set ServerDatacenter=74YFP-3QFB3-KQT8W-PMXWJ-7M648
set ServerEnterpriseIA64=GT63C-RJFQ3-4GMB6-BRFB9-CB83V
goto winact

:win8
echo 当前系统为Windows 8/2012。
set Professional=NG4HW-VH26C-733KW-K6F98-J8CK4
set ProfessionalN=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
set Core=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
set Enterprise=32JNW-9KQ84-P47T8-D8GGY-CWCK7
set EnterpriseN=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
set CoreN=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
set CoreSingleLanguage=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
set CoreCountrySpecific=4K36P-JN4VD-GDC6V-KDT89-DYFKP
set ServerMultiPointPremium=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
set ServerMultiPointStandard=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
set ServerStandard=XC9B7-NBPP2-83J2H-RHMBY-92BT4
set ServerDatacenter=48HP8-DN98B-MYWDG-T2DCC-8W83P
goto winact

:win81
echo 当前系统为Windows 8.1。
set Professional=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
set ProfessionalN=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
set Enterprise=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
set EnterpriseN=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
set ServerSolution=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
set ServerStandard=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerDatacenter=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set EmbeddedIndustry=32JNW-9KQ84-P47T8-D8GGY-CWCK7
goto winact

:win10
echo 当前系统为Windows 10/Server 2016-2019
echo.&echo 尝试进行永久激活……
cscript //Nologo %windir%\system32\slmgr.vbs /ipk VK7JG-NPHTM-C97JM-9MPGT-3V66T >nul
cscript //Nologo %windir%\system32\slmgr.vbs /ato | find "成功" >nul && goto win10done
cscript //Nologo %windir%\system32\slmgr.vbs /ipk QJNXR-YD97Q-K7WH4-RYWQ8-6MT6Y >nul
cscript //Nologo %windir%\system32\slmgr.vbs /ato | find "成功" >nul && goto win10done
echo 永久激活失败QwQ，尝试进行KMS激活……
echo.
set Core=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
set CoreCountrySpecific=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
set CoreN=3KHY7-WNT83-DGQKR-F7HPR-844BM
set CoreSingleLanguage=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
set Professional=W269N-WFGWX-YVC9B-4J6C9-T83GX
set ProfessionalN=MH37W-N47XK-V7XM9-C7227-GCQG9
set Enterprise=NPPR9-FWDCX-D2C8J-H872K-2YT43
set EnterpriseN=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
set Education=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
set EducationN=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
set EnterpriseS=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
set EnterpriseSN=2F77B-TNFGY-69QQF-B8YKP-D69TJ
set ServerDatacenter=CB7KF-BWN84-R7R2Y-793K2-8XDDG
set ServerStandard=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
set ServerEssentials=JCKRF-N37P4-C2D82-9YXRT-4M63B
set EnterpriseG=YYVX9-NTFWV-6MDM3-9PT4T-4M68B
set EnterpriseGN=44RPN-FTY23-9VTTB-MP9BX-T84FV
set ProfessionalEducation=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
set ProfessionalEducationN=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
set ProfessionalWorkstation=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstations=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstationsN=9FNHH-K3HBT-3W4TD-6383H-6XYWF
set ServerDatacenter=WMDGN-G9PQG-XVVXX-R3X43-63DFG
set ServerStandard=N69G4-B89J2-4G8F4-WWYCC-J464C
set ServerEssentials=WVDHN-86M7X-466P6-VHXV7-YY726
set ServerRdsh=CPWHC-NT2C7-VYW78-DHDB2-PG3GK
goto winact

:winact
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%i
if defined %EditionID% (
	cscript //Nologo %windir%\system32\slmgr.vbs /ipk !%EditionID%! >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /skms %KMSSERVER% >nul
	cscript //Nologo %windir%\system32\slmgr.vbs /ato
) else (
	echo 找不到序列号，可能是旗舰版之类的系统QwQ
)
echo.&pause
goto begin

:win10done
echo.&echo Windows已经永久激活！
goto begin

:choose2a
echo 检查安装的Office……
call :GetOfficePath 14 Office2010
call :ActOffice 14 Office2010
call :GetOfficePath 15 Office2013
call :ActOffice 15 Office2013
call :GetOfficePath 16 Office2016
call :ActOffice 16 Office2016
echo.&echo 检查是否存在office 2016零售版，将执行Retail转Vol
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles(x86)%\Microsoft Office\Office16
if DEFINED _Office16Path (echo 已发现 Office2016 零售版
    call :ActOffice 16 Office2016
  ) else (echo 未发现 Office2016零售版)
echo 操作完成，按任意键返回…… & pause>nul
goto begin

:ActOffice
if DEFINED _Office%1Path (
    cd /d "!_Office%1Path!"
	echo.&echo 检查 %2 的激活状态……
	cscript //nologo ospp.vbs /act | find /i "successful" > nul && (
        echo %2 已经激活，自动跳过 & echo.) || (
		echo %2 未激活，正尝试进行激活……
		if %1 EQU 16 call :Licens16
		cscript //nologo ospp.vbs /sethst:%KMSSERVER% >nul
		cscript //nologo ospp.vbs /act | find /i "successful" && (
        echo.&echo ***** %2 激活成功 ***** & echo.) || (echo.&echo ***** %2 激活失败 ***** & echo.)
		)
)    
cd /d "%~dp0"
goto :EOF

:GetOfficePath
echo.&echo 正在检测 %2 产品的安装路径……
set _Office%1Path=
set _Reg32=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot
set _Reg64=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot
reg query "%_Reg32%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg32%" /v "Path"') do SET "_OfficePath1=%%b"
reg query "%_Reg64%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg64%" /v "Path"') do SET "_OfficePath2=%%b"
if DEFINED _OfficePath1 (if exist "%_OfficePath1%ospp.vbs" set _Office%1Path=!_OfficePath1!)
if DEFINED _OfficePath2 (if exist "%_OfficePath2%ospp.vbs" set _Office%1Path=!_OfficePath2!)
set _OfficePath1=
set _OfficePath2=
if DEFINED _Office%1Path (echo 已发现 %2) else (echo 未发现 %2)
goto :EOF

:Licens16
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul
cscript ospp.vbs /inpkey:NYH39-6GMXT-T39D4-WVXY2-D69YY >nul
cscript ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4 >nul
cscript ospp.vbs /inpkey:RBWW7-NTJD4-FFK2C-TDJ7V-4C2QP >nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT >nul
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul
goto :EOF
goto begin

:choose3
echo ---------------------------------------------------
echo 果核剥壳的备用KMS服务器为
echo.
echo kms.qkeke.com
echo.
echo 当然，你可以输入其他任何你想使用的KMS服务器地址,如
echo kms.loli.beer            kms.cin.ink
echo kms.loli.cab             kms.izetn.cn
echo kms.loli.best            kms.ddddg.cn
echo kms.iaini.net            kms.zhuxiaole.org
echo kms.kuretru.com          kms.magicwall.org
echo kms.03k.org              kms.luody.info
echo kms.v0v.bid              kms.bige0.com
echo kms.ddz.red              windows.kms.app
echo kms.lotro.cc             cy2617.jios.org
echo kms.ijio.net             enter.picp.net
echo kms.heng07.com           xykz.f3322.org
echo kms.chinancce.com        nb.shenqw.win
echo kms8.MSGuides.com        zh.us.to
echo kms.moeclub.org          home.aalook.com
echo kms.cangshui.net         kms.kmzs123.cn
echo ---------------------------------------------------
echo 请输入激活Windows的KMS服务器地址(回车默认kms.qkeke.com)：
echo 你可以使用其他激活服务器，但不检查其有效性
set KMSSERVER=kms.qkeke.com
set /p KMSSERVER=
goto choose1a

:choose4
echo ---------------------------------------------------
echo 果核剥壳的备用KMS服务器为
echo.
echo kms.qkeke.com
echo.
echo 当然，你可以输入其他任何你想使用的KMS服务器地址,如
echo kms.loli.beer            kms.cin.ink
echo kms.loli.cab             kms.izetn.cn
echo kms.loli.best            kms.ddddg.cn
echo kms.iaini.net            kms.zhuxiaole.org
echo kms.kuretru.com          kms.magicwall.org
echo kms.03k.org              kms.luody.info
echo kms.v0v.bid              kms.bige0.com
echo kms.ddz.red              windows.kms.app
echo kms.lotro.cc             cy2617.jios.org
echo kms.ijio.net             enter.picp.net
echo kms.heng07.com           xykz.f3322.org
echo kms.chinancce.com        nb.shenqw.win
echo kms8.MSGuides.com        zh.us.to
echo kms.moeclub.org          home.aalook.com
echo kms.cangshui.net         kms.kmzs123.cn
echo ---------------------------------------------------
echo 请输入激活Office的KMS服务器地址(回车默认kms.qkeke.com)：
echo 你可以使用其他激活服务器，但不检查其有效性
set KMSSERVER=kms.qkeke.com
set /p KMSSERVER=
goto choose2a

:choose5
set /p choose=真的要清除Windows的KMS激活？【1】继续   【2】关闭：
if /i "%choose%"=="1" goto unactw
if /i "%choose%"=="2" goto begin
:unactw
slmgr /upk
slmgr /ckms
slmgr /rearm
cls
echo 清除完成awa，请重启电脑
pause
goto begin

:choose6
set /p choose=真的要清除Office的KMS激活？【1】继续   【2】关闭：
if /i "%choose%"=="1" goto unacto
if /i "%choose%"=="2" goto begin
:unacto
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles%\Microsoft Office\Office14\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:6MWKP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:D69YY
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:W8GF4
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:4C2QP
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:WFG99
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:G83KT
cscript "%ProgramFiles(x86)%\Microsoft Office\Office14\OSPP.VBS" /unpkey:RJRJK
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /remhst
cls
echo 清除完成awa
pause
goto begin

:choose7
echo 正在查询Windows激活状态……
cscript %windir%\system32\slmgr.vbs -xpr
cscript %windir%\system32\slmgr.vbs -dlv
pause
goto begin

:choose8
echo -----请选择你的Office版本-----
echo [1]Office 2016/2019 x64
echo [2]Office 2016/2019 x86
echo [3]Office 2013 x64
echo [4]Office 2013 x86
echo [5]Office 2010 x64
echo [6]Office 2010 x86
echo [7]手动输入Office目录
echo --------请输入你的选择--------
set /p OFFICE=
if /i "%OFFICE%"=="1" cls&set OPATH=%ProgramFiles%\Microsoft Office\Office16
if /i "%OFFICE%"=="2" cls&set OPATH=%ProgramFiles(x86)%\Microsoft Office\Office16
if /i "%OFFICE%"=="3" cls&set OPATH=%ProgramFiles%\Microsoft Office\Office15
if /i "%OFFICE%"=="4" cls&set OPATH=%ProgramFiles(x86)%\Microsoft Office\Office15
if /i "%OFFICE%"=="5" cls&set OPATH=%ProgramFiles%\Microsoft Office\Office14
if /i "%OFFICE%"=="6" cls&set OPATH=%ProgramFiles(x86)%\Microsoft Office\Office14
if /i "%OFFICE%"=="7" cls&goto inputopath
goto checkoact

:inputopath
echo 请输入你的Office路径：
echo 如C:\Program Files\Microsoft Office\Office16
set /p OPATH=
goto checkoact

:checkoact
echo 你当前的Office目录为%OPATH%
echo.
cscript "%OPATH%\OSPP.VBS" /dstatus
pause
goto begin

:choose9
start http://edragon.sxl.cn
goto begin

:choose0
echo ———发布链接表【Page 1/1】———
echo [1]github项目帖
echo [2]果核剥壳发布贴
echo ———请选择你要访问的网址：———
set /p WEB=
if /i "%WEB%"=="1" cls&start https://github.com/EDragon007/EDragonKMS
if /i "%WEB%"=="2" cls&start https://www.ghpym.com/ghkmsbat.html
goto begin