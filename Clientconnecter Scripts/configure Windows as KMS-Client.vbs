
'KMS-Client-Keys: http://technet.microsoft.com/en-us/library/jj612867.aspx
'KMS Client Keys: http://www.infrastrukturhelden.de/microsoft-infrastruktur/key-management-service-kms-client-seriennummern-updated.html

' find KMS Server via DNS SRV Record
autodetect="True"
' or connect to this specified KMS Server
kms_server="kms.walhalla.local"
kms_port="1688"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
varos = Trim(objOperatingSystem.Caption)
'Msgbox objOperatingSystem.Version, 0 + 32,"Window Version"
Next

Select Case varos
'Windows 10 ----------------------------
    Case "Windows 10 Professional"
		kms_serial = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
    Case "Windows 10 Professional N"
		kms_serial = "MH37W-N47XK-V7XM9-C7227-GCQG9"
    Case "Windows 10 Enterprise"
		kms_serial = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    Case "Windows 10 Enterprise N"
		kms_serial = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
    Case "Windows 10 Education"
		kms_serial = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
    Case "Windows 10 Education N"
		kms_serial = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
    Case "Windows 10 Enterprise 2015 LTSB"
		kms_serial = "WNMTR-4C88C-JK8YV-HQ7T2-76DF9"
    Case "Windows 10 Enterprise 2015 LTSB N"
		kms_serial = "2F77B-TNFGY-69QQF-B8YKP-D69TJ"
 
'Windows 8.1 ----------------------------
    Case "Microsoft Windows 8.1 Professional"
		kms_serial = "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9"
	Case "Microsoft Windows 8.1 Professional N"
		kms_serial = "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY"
	Case "Microsoft Windows 8.1 Enterprise"
		kms_serial = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7"
	Case "Microsoft Windows 8.1 Enterprise N"
		kms_serial = "TT4HM-HN7YT-62K67-RGRQJ-JFFXW"

'Windows 8 ------------------------------
    Case "Microsoft Windows 8 Professional"
		kms_serial = "NG4HW-VH26C-733KW-K6F98-J8CK4"
    Case "Microsoft Windows 8 Professional N"
		kms_serial = "XCVCF-2NXM9-723PB-MHCB7-2RYQQ"
    Case "Microsoft Windows 8 Enterprise"
		kms_serial = "32JNW-9KQ84-P47T8-D8GGY-CWCK7"
	Case "Microsoft Windows 8 Enterprise N"
		kms_serial = "JMNMF-RHW7P-DMY6X-RF3DR-X2BQT"

'Windows 7 ------------------------------
    Case "Microsoft Windows 7 Professional"
		kms_serial = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"
    Case "Microsoft Windows 7 Professional N"
		kms_serial = "MRPKT-YTG23-K7D7T-X2JMM-QY7MG"
    Case "Microsoft Windows 7 Professional E"
		kms_serial = "W82YF-2Q76Y-63HXB-FGJG9-GF7QX"
    Case "Microsoft Windows 7 Enterprise"
		kms_serial = "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"
	Case "Microsoft Windows 7 Enterprise N"
		kms_serial = "YDRBP-3D83W-TY26F-D46B2-XCKRJ"
	Case "Microsoft Windows 7 Enterprise E"
		kms_serial = "C29WB-22CC8-VJ326-GHFJW-H9DH4"

'Windows Vista ------------------------------
    Case "Microsoft Windows Vista Business"
		kms_serial = "YFKBB-PQJJV-G996G-VWGXY-2V3X8"
    Case "Microsoft Windows Vista Business N"
		kms_serial = "HMBQG-8H2RH-C77VX-27R82-VMQBT"
    Case "Microsoft Windows Vista Enterprise"
		kms_serial = "VKK3X-68KWM-X2YGT-QR4M6-4BWMV"
    Case "Microsoft Windows Vista Enterprise N"
		kms_serial = "VTC42-BM838-43QHV-84HX6-XJXKV"

'Windows Server 2012 R2 ---------------------------
	Case "Microsoft Windows Server 2012 R2 Standard"
		kms_serial = "D2N9P-3P6X9-2R39C-7RTCD-MDVJX"
	Case "Microsoft Windows Server 2012 R2 Datacenter"
		kms_serial = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"
	Case "Microsoft Windows Windows Server 2012 R2 Essentials"
		kms_serial = "KNC87-3J2TX-XB4WP-VCPJV-M4FWM"

'Windows Server 2012 ------------------------------
	Case "Microsoft Windows Server 2012 Core"
		kms_serial = "BN3D2-R7TKB-3YPBD-8DRP2-27GG4"
	Case "Microsoft Windows Server 2012 Core N"
		kms_serial = "8N2M2-HWPGY-7PGT9-HGDD8-GVGGY"
	Case "Microsoft Windows Server 2012 Core Single Language"
		kms_serial = "2WN2H-YGCQR-KFX6K-CD6TF-84YXQ"
	Case "Microsoft Windows Server 2012 Core Country Specific"
		kms_serial = "4K36P-JN4VD-GDC6V-KDT89-DYFKP"
	Case "Microsoft Windows Server 2012 Server Standard"
		kms_serial = "XC9B7-NBPP2-83J2H-RHMBY-92BT4"
	Case "Microsoft Windows Server 2012 Standard Core"
		kms_serial = "XC9B7-NBPP2-83J2H-RHMBY-92BT4"
	Case "Microsoft Windows Server 2012 Multipoint Standard"
		kms_serial = "HM7DN-YVMH3-46JC3-XYTG7-CYQJJ"
	Case "Microsoft Windows Server 2012 Multipoint Premium"
		kms_serial = "XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G"
	Case "Microsoft Windows Server 2012 Datacenter"
		kms_serial = "48HP8-DN98B-MYWDG-T2DCC-8W83P"
	Case "Microsoft Windows Server 2012 Datacenter Core"
		kms_serial = "48HP8-DN98B-MYWDG-T2DCC-8W83P"

'Windows Server 2008 R2 ------------------------------
	Case "Microsoft Windows Server 2008 R2 Web"
		kms_serial = "6TPJF-RBVHG-WBW2R-86QPH-6RTM4"
	Case "Microsoft Windows Server 2008 HPC edition"
		kms_serial = "TT8MH-CG224-D3D7Q-498W2-9QCTX"
	Case "Microsoft Windows Server 2008 R2 Standard"
		kms_serial = "YC6KT-GKW9T-YTKYR-T4X34-R7VHC"
	Case "Microsoft Windows Server 2008 R2 Enterprise"
		kms_serial = "489J6-VHDMP-X63PK-3K798-CPX3Y"
	Case "Microsoft Windows Server 2008 R2 Datacenter"
		kms_serial = "74YFP-3QFB3-KQT8W-PMXWJ-7M648"
	Case "Microsoft Windows Server 2008 R2 for Itanium-based Systems"
		kms_serial = "GT63C-RJFQ3-4GMB6-BRFB9-CB83V"

'Windows Server 2008 ------------------------------
    Case "Microsoft Windows Web Server 2008"
		kms_serial = "WYR28-R7TFJ-3X2YQ-YCY4H-M249D"
    Case "Microsoft Windows Server 2008 Standard"
		kms_serial = "TM24T-X9RMF-VWXK6-X8JC9-BFGM2"
    Case "Microsoft Windows Server 2008 Standard without Hyper-V"
		kms_serial = "W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ"
    Case "Microsoft Windows Server 2008 Enterprise"
		kms_serial = "YQGMW-MPWTJ-34KDK-48M3W-X4Q6V"
    Case "Microsoft Windows Server 2008 Enterprise without Hyper-V"
		kms_serial = "39BXF-X8Q23-P2WWT-38T2F-G3FPG"
    Case "Microsoft Windows Server 2008 HPC"
		kms_serial = "RCTX3-KWVHP-BR6TB-RB6DM-6X7HP"
    Case "Microsoft Windows Server 2008 Datacenter"
		kms_serial = "7M67G-PC374-GR742-YH8V4-TCBY3"
    Case "Microsoft Windows Server 2008 Datacenter without Hyper-V"
		kms_serial = "22XQ2-VRXRG-P8D42-K34TD-G3QQC"
    Case "Microsoft Windows Server 2008 for Itanium-Based Systems"
		kms_serial = "4DWFP-JF3DJ-B7DTH-78FJB-PDRHK"

    Case Else
		Msgbox ("Abbruch: kein unterstütztes Betriebssystem gefunden! ->" + varos)
		WScript.Quit
End Select


set shell = CreateObject("WScript.Shell")

'set serial to KMS-Client
shell.run "slmgr.vbs /ipk " + kms_serial,1,True
if autodetect = "True" Then
	Msgbox ("setze KMS Server auf Autodetect ...")
	shell.run "slmgr.vbs /ckms",1,True
else
	Msgbox ("schreibe Adresse des KMS Server in die Registry ...")
	'set KMS-Server IP:Port
	shell.run "slmgr.vbs /skms " + kms_server + ":" + kms_port,1,True
End if
'active Windows
shell.run "slmgr.vbs /ato ",1,True
