Attribute VB_Name = "ConnectionProxy"
Option Explicit
' Author: Mohammad Maysami
' Usage: Detecting Various Proxy Configurations (Auto Detect, Auto Config URL PAC, Proxy ...)
'   GetProxyInfoForUrl(Optional URL, Optional ProxyDetails As Variant) As ProxyInfo
'   	Syntax1: GetProxyInfoForUrl()
'   	Syntax2: GetProxyInfoForUrl("http://www.google.com", ProxyDetails)
'   	Syntax3: GetProxyInfoForUrl(Array("http://www.google.com", "http://www.microsoft.com"), ProxyDetails)
'
'
'   Possible AutoProxy Errors:
'       12166 - error in proxy auto-config script code
'       12167 - unable to download proxy auto-config script
'       12180 - WPAD detection failed
'
'   Adapted from Stephen Sulzer 2004

'=============================================================
'                   Type Structure Definitions
'=============================================================

'--------------------------------------
'           My ProxyInfo
'--------------------------------------
' Type Structure for my Connection Proxy Information
Public Type ProxyInfo
   ProxyActive As Boolean
   ProxyServer As String
   ProxyBypass As String
End Type

#If VBA7 Then
    '--------------------------------------
    '           IE PROXY CONFIG
    '--------------------------------------
    ' Type Structure for IE Proxy Settings
    Private Type WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
       fAutoDetect As Long
       lpszAutoConfigUrl As LongPtr
       lpszProxy As LongPtr
       lpszProxyBypass As LongPtr
    End Type
    
    '--------------------------------------
    '           WinHttp Proxy Info
    '--------------------------------------
    Private Type WINHTTP_PROXY_INFO
       dwAccessType As Long
       lpszProxy As LongPtr
       lpszProxyBypass As LongPtr
    End Type
    
    '--------------------------------------
    '           AutoProxy Options
    '--------------------------------------
    ' Type Structure for AutoProxy Options
    Private Type WINHTTP_AUTOPROXY_OPTIONS
       dwFlags As Long
       dwAutoDetectFlags As Long
       lpszAutoConfigUrl As LongPtr
       lpvReserved As LongPtr
       dwReserved As Long
       fAutoLogonIfChallenged As Long
    End Type

#Else
    '--------------------------------------
    '           IE PROXY CONFIG
    '--------------------------------------
    ' Type Structure for IE Proxy Settings
    Private Type WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
       fAutoDetect As Long
       lpszAutoConfigUrl As Long
       lpszProxy As Long
       lpszProxyBypass As Long
    End Type
    
    '--------------------------------------
    '           WinHttp Proxy Info
    '--------------------------------------
    Private Type WINHTTP_PROXY_INFO
       dwAccessType As Long
       lpszProxy As Long
       lpszProxyBypass As Long
    End Type
    
    '--------------------------------------
    '           AutoProxy Options
    '--------------------------------------
    ' Type Structure for AutoProxy Options
    Private Type WINHTTP_AUTOPROXY_OPTIONS
       dwFlags As Long
       dwAutoDetectFlags As Long
       lpszAutoConfigUrl As Long
       lpvReserved As Long
       dwReserved As Long
       fAutoLogonIfChallenged As Long
    End Type
#End If

' AutoProxy Options Constants
'--------------------------------------
' Constants for dwFlags of WINHTTP_AUTOPROXY_OPTIONS
Private Const WINHTTP_AUTOPROXY_AUTO_DETECT = 1
Private Const WINHTTP_AUTOPROXY_CONFIG_URL = 2
 
' Constants for dwAutoDetectFlags of WINHTTP_AUTOPROXY_OPTIONS
Private Const WINHTTP_AUTO_DETECT_TYPE_DHCP = 1
Private Const WINHTTP_AUTO_DETECT_TYPE_DNS = 2

' Constants for URLs to Ping and AutoDetect Proxy
Private Const NRConnectionURL1 As String = "http://www.microsoft.com"
Private Const NRConnectionURL2 As String = "http://www.google.com"
Private Const NRConnectionURL3 As String = "http://www.wikipedia.com"
'=============================================================
'                     Lib Declarations
'=============================================================
' VBA7 IF To Address both 32/64-bits
#If VBA7 Then
    ' Need CopyMemory to copy BSTR pointers around
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, _
         ByVal lpSource As LongPtr, ByVal cbCopy As Long)
         
    ' SysAllocString creates a UNICODE BSTR string based on a UNICODE string
    Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal pwsz As LongPtr) As LongPtr
    
    ' Need GlobalFree to free the pointers in the CURRENT_USER_IE_PROXY_CONFIG
    ' structure returned from WinHttpGetIEProxyConfigForCurrentUser, per the documentation
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal p As LongPtr) As LongPtr
    
    ' https://docs.microsoft.com/en-us/windows/desktop/api/winhttp/nf-winhttp-winhttpgetieproxyconfigforcurrentuser
    'BOOLAPI WinHttpGetIEProxyConfigForCurrentUser(
    '  IN OUT WINHTTP_CURRENT_USER_IE_PROXY_CONFIG *pProxyConfig);
    Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" _
       (ByRef proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long
       
    ' https://docs.microsoft.com/en-us/windows/desktop/api/winhttp/nf-winhttp-winhttpgetproxyforurl
    ' Returns 0 on Fail, Number Otherwise ?
    ' BOOLAPI WinHttpGetProxyForUrl(
    '  IN HINTERNET                 hSession,
    '  IN LPCWSTR                   lpcwszUrl,
    '  IN WINHTTP_AUTOPROXY_OPTIONS *pAutoProxyOptions,
    '  OUT WINHTTP_PROXY_INFO       *pProxyInfo);
    Private Declare PtrSafe Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" _
       (ByVal hSession As LongPtr, _
        ByVal pszUrl As LongPtr, _
        ByRef pAutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS, _
        ByRef pProxyInfo As WINHTTP_PROXY_INFO) As Long
     
    ' https://docs.microsoft.com/en-us/windows/desktop/api/winhttp/nf-winhttp-winhttpopen
    'WINHTTPAPI HINTERNET WinHttpOpen(
    '  LPCWSTR pszAgentW,
    '  DWORD   dwAccessType,
    '  LPCWSTR pszProxyW,
    '  LPCWSTR pszProxyBypassW,
    '  DWORD   dwFlags);
    Private Declare PtrSafe Function WinHttpOpen Lib "WinHTTP.dll" _
       (ByVal pszUserAgent As LongPtr, _
        ByVal dwAccessType As Long, _
        ByVal pszProxyName As LongPtr, _
        ByVal pszProxyBypass As LongPtr, _
        ByVal dwFlags As Long) As LongPtr
     
    ' https://docs.microsoft.com/en-us/windows/desktop/api/winhttp/nf-winhttp-winhttpclosehandle
    ' BOOLAPI WinHttpCloseHandle(
    '  IN HINTERNET hInternet);
    Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHTTP.dll" _
       (ByVal hInternet As LongPtr) As Long
	 

#Else
    Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal lpDest As Long, _
         ByVal lpSource As Long, ByVal cbCopy As Long)
         
    Private Declare Function SysAllocString Lib "oleaut32" (ByVal pwsz As Long) As Long
    
    Private Declare Function GlobalFree Lib "kernel32" (ByVal p As Long) As Long
    
    Private Declare Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" _
       (ByRef proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long
       
    Private Declare Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" _
       (ByVal hSession As Long, _
        ByVal pszUrl As Long, _
        ByRef pAutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS, _
        ByRef pProxyInfo As WINHTTP_PROXY_INFO) As Long
     
    Private Declare Function WinHttpOpen Lib "WinHTTP.dll" _
       (ByVal pszUserAgent As Long, _
        ByVal dwAccessType As Long, _
        ByVal pszProxyName As Long, _
        ByVal pszProxyBypass As Long, _
        ByVal dwFlags As Long) As Long
     
    Private Declare Function WinHttpCloseHandle Lib "WinHTTP.dll" _
       (ByVal hInternet As Long) As Long
#End If

   

'**************************************************************************************************************************
'**************************************************************************************************************************
'**************************************************************************************************************************

'=============================================================
'                   Get Proxy Info
'=============================================================
 
Public Function GetProxyInfoForUrl(Optional URL, Optional ProxyDetails As Variant) As ProxyInfo
    ' Using a user-defined Full (Array of) URL(s), Get IE Proxy Config and
    '   Find Proxy based on Auto Detect Protocols (AutoDetect, AutoConfigURL for PAC File)
    '   or by IE Proxy if available
    '   It returns ProxyInfo Structure (Boolean Active,String Proxy, String ProxyByPass
    '
    '   Syntax1: GetProxyInfoForUrl()
    '   Syntax2: GetProxyInfoForUrl("http://www.google.com", ProxyDetails)
    '   Syntax3: GetProxyInfoForUrl(Array("http://www.google.com", "http://www.microsoft.com"), ProxyDetails)
    '
    '   Inputs:
    '       opt IN  URL(s)      : Array of or Single String Full URLs to AutoDetect Proxy
    '       opt OUT ProxyDetails: Custom IE Proxy Structure to Pass out IE Proxy Details and Status Code
    '                (1) = IE AutoDetect    (fAutoDetect)
    '                (2) = IE AutoCofigUrl  (lpszAutoConfigUrl)
    '                (3) = IE Proxy         (lpszProxy)
    '                (4) = IE Proxy Bypass  (lpszProxyBypass)
    '                (5) = DevCode
    '
    '   Outputs:
    '       OUT ProxyInfo      : Boolean Active, String Proxy, String ProxyBypass
    '
    '   Notes, Possible AutoProxy Errors:
    '       12166 - error in proxy auto-config script code
    '       12167 - unable to download proxy auto-config script
    '       12180 - WPAD detection failed
    '
    Dim IEProxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    Dim AutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS
    Dim WinHttpProxyInfo As WINHTTP_PROXY_INFO
    Dim ProxyInfo As ProxyInfo
    
    'Dim fStatusProxy As Integer
    Dim fDoAutoProxy As Boolean
    #If VBA7 Then
        Dim ProxyStringPtr As LongPtr
        Dim ptr As LongPtr
    #Else
        Dim ProxyStringPtr As Long
        Dim ptr As Long
    #End If
    Dim error As Long
    Dim DevCode As String
    Dim trial As Integer
    Dim MaxTrial As Integer
    
    ' --------------------------------------------
    ' Init. URLs and Max Trials
    ' --------------------------------------------
    If IsMissing(URL) Then
        URL = Array(NRConnectionURL1)
        MaxTrial = 1
    Else
        '
        If IsArray(URL) Then
            MaxTrial = UBound(URL) - LBound(URL) + 1
        ElseIf WorksheetFunction.IsText(URL) Then
            URL = Array(URL)
            MaxTrial = 1
        Else
            URL = Array(NRConnectionURL1)
            MaxTrial = 1
        End If
    End If
    
    ' --------------------------------------------
    ' Reset/Init Class Instances
    ' --------------------------------------------
    ' Init ProxyInfo
    ProxyInfo.ProxyActive = False
    ProxyInfo.ProxyServer = vbNullString
    ProxyInfo.ProxyBypass = vbNullString
        
    ' Init WinHttpProxyInfo
    WinHttpProxyInfo.dwAccessType = 0
    WinHttpProxyInfo.lpszProxy = 0
    WinHttpProxyInfo.lpszProxyBypass = 0
    
    ' Init IEProxyConfig
    IEProxyConfig.fAutoDetect = 0
    IEProxyConfig.lpszAutoConfigUrl = 0
    IEProxyConfig.lpszProxy = 0
    IEProxyConfig.lpszProxyBypass = 0
    
    ' Init AutoProxyOptions
    AutoProxyOptions.dwFlags = 0
    AutoProxyOptions.dwAutoDetectFlags = 0
    AutoProxyOptions.lpszAutoConfigUrl = 0
    AutoProxyOptions.dwReserved = 0
    AutoProxyOptions.lpvReserved = 0
    AutoProxyOptions.fAutoLogonIfChallenged = 1
    

    ' Other Flags
    'fStatusProxy = 0
    fDoAutoProxy = False
    ProxyStringPtr = 0
    ptr = 0
    DevCode = ""
    trial = 0

    ' --------------------------------------------
    ' Check IE's proxy configuration
    ' --------------------------------------------
    If (WinHttpGetIEProxyConfigForCurrentUser(IEProxyConfig) > 0) Then
        ' If IE is configured to auto-detect, then we will too.
        If (IEProxyConfig.fAutoDetect <> 0) Then
            'fStatusProxy = fStatusProxy + 1
            DevCode = DevCode & vbCrLf & "[IE Auto Detect]"
            AutoProxyOptions.dwFlags = WINHTTP_AUTOPROXY_AUTO_DETECT
            AutoProxyOptions.dwAutoDetectFlags = _
                        WINHTTP_AUTO_DETECT_TYPE_DHCP + _
                        WINHTTP_AUTO_DETECT_TYPE_DNS
            fDoAutoProxy = True
        End If
    
        ' If IE is configured to use an auto-config script, then
        ' we will use it too
        If (IEProxyConfig.lpszAutoConfigUrl <> 0) Then
            'fStatusProxy = fStatusProxy + 10
            DevCode = DevCode & vbCrLf & "[AutoConfigUrl PAC]"
			
            AutoProxyOptions.dwFlags = AutoProxyOptions.dwFlags + _
                        WINHTTP_AUTOPROXY_CONFIG_URL
                        
            'If dwFlags includes the WINHTTP_AUTOPROXY_CONFIG_URL flag,
            '   the lpszAutoConfigUrl must point to a null-terminated Unicode string
            '   that contains the URL of the proxy auto-configuration (PAC) file.
            AutoProxyOptions.lpszAutoConfigUrl = IEProxyConfig.lpszAutoConfigUrl
            
            fDoAutoProxy = True
        End If
        
    Else
        'fStatusProxy = fStatusProxy + 100
        DevCode = DevCode & vbCrLf & "[No Proxy Config]"
		
        ' if the IE proxy config is not available, then
        ' we will try auto-detection
        AutoProxyOptions.dwFlags = WINHTTP_AUTOPROXY_AUTO_DETECT
        AutoProxyOptions.dwAutoDetectFlags = _
                        WINHTTP_AUTO_DETECT_TYPE_DHCP + _
                        WINHTTP_AUTO_DETECT_TYPE_DNS
        fDoAutoProxy = True
    End If
    
   
    ' --------------------------------------------
    '   Handle Auto Proxy Configurations
    ' --------------------------------------------
    If fDoAutoProxy Then
        #If VBA7 Then
            Dim hSession As LongPtr
        #Else
            Dim hSession As Long
        #End If
        

        ' Need to create a temporary WinHttp session handle
        '  Note: performance of this GetProxyInfoForUrl function can be
        '   improved by saving this hSession handle across calls
        '   instead of creating a new handle each time
        hSession = WinHttpOpen(0, 1, 0, 0, 0)
    
        Do While trial < MaxTrial
            trial = trial + 1
            If (WinHttpGetProxyForUrl(hSession, StrPtr(URL(trial - 1)), AutoProxyOptions, _
                    WinHttpProxyInfo) > 0) Then
                DevCode = DevCode & vbCrLf & "{Pass" & trial & ": " & WinHttpProxyInfo.lpszProxy & "}"
                ProxyStringPtr = WinHttpProxyInfo.lpszProxy
                ' Ignore WinHttpProxyInfo.lpszProxyBypass, it will not be set
                If (ProxyStringPtr <> 0) Then
                    ' Terminate Trial Loop if Found
                    trial = MaxTrial + 1
                End If
            Else
                ' some possibly autoproxy errors:
                '   12166 - error in proxy auto-config script code
                '   12167 - unable to download proxy auto-config script
                '   12180 - WPAD detection failed
                error = Err.LastDllError
                Select Case error
                    Case 12166
                        DevCode = DevCode & vbCrLf & "{Fail" & trial & ": PAC Script Execution}"
                    Case 12167
                        DevCode = DevCode & vbCrLf & "{Fail" & trial & ": PAC File Download}"
                    Case 12180
                        DevCode = DevCode & vbCrLf & "{Fail" & trial & ": PAC URL (WPAD) Detection}"
                    Case Else
                        DevCode = DevCode & vbCrLf & "{Fail" & trial & ": " & error & "}"
                End Select

									
											 
															  
																   
											 
            End If
        Loop
        WinHttpCloseHandle (hSession)
    End If
    
    
    ' --------------------------------------------
    ' Check IE Proxy, If NO Proxy Detected
    ' --------------------------------------------
    ' If we don't have a proxy server from WinHttpGetProxyForUrl,
    ' then pick one up from the IE proxy config (if given)
    If (ProxyStringPtr = 0) Then
        DevCode = DevCode & vbCrLf & "[Empty ProxyForUrl String]"
        ProxyStringPtr = IEProxyConfig.lpszProxy
    End If
    
    
    ' --------------------------------------------
    ' Convert Proxy to Basic Strings ==> ProxyInfo
    ' --------------------------------------------
    ' If there's a proxy string, convert it to a Basic string
    If (ProxyStringPtr <> 0) Then
        'fStatusProxy = fStatusProxy + 1000
        DevCode = DevCode & vbCrLf & "[IE Proxy Config]"
        ptr = SysAllocString(ProxyStringPtr)
        CopyMemory VarPtr(ProxyInfo.ProxyServer), VarPtr(ptr), 4
        ProxyInfo.ProxyActive = True
    End If
    
    
    ' --------------------------------------------
    '  Pick IE Proxy ByPass ==> ProxyInfo
    ' --------------------------------------------
    ' Pick up any bypass string from the IEProxyConfig
    If (IEProxyConfig.lpszProxyBypass <> 0) Then
        ptr = SysAllocString(IEProxyConfig.lpszProxyBypass)
        CopyMemory VarPtr(ProxyInfo.ProxyBypass), VarPtr(ptr), 4
    End If
    
       
    If Not IsMissing(ProxyDetails) Then
        ReDim ProxyDetails(5) As Variant
        ProxyDetails(1) = IEProxyConfig.fAutoDetect
        ProxyDetails(2) = IEProxyConfig.lpszAutoConfigUrl
        ProxyDetails(3) = IEProxyConfig.lpszProxy
        ProxyDetails(4) = IEProxyConfig.lpszProxyBypass
        ProxyDetails(5) = DevCode
    End If
    
    GetProxyInfoForUrl = ProxyInfo
    
    ' --------------------------------------------
    ' Free Up Memory/Pointers
    ' --------------------------------------------
    ' Free any strings received from WinHttp APIs
    If (IEProxyConfig.lpszAutoConfigUrl <> 0) Then
        GlobalFree (IEProxyConfig.lpszAutoConfigUrl)
    End If
    If (IEProxyConfig.lpszProxy <> 0) Then
        GlobalFree (IEProxyConfig.lpszProxy)
    End If
    If (IEProxyConfig.lpszProxyBypass <> 0) Then
        GlobalFree (IEProxyConfig.lpszProxyBypass)
    End If
    If (WinHttpProxyInfo.lpszProxy <> 0) Then
        GlobalFree (WinHttpProxyInfo.lpszProxy)
    End If
    If (WinHttpProxyInfo.lpszProxyBypass <> 0) Then
        GlobalFree (WinHttpProxyInfo.lpszProxyBypass)
    End If
           
End Function


'**************************************************************************************************************************
'**************************************************************************************************************************
'**************************************************************************************************************************
