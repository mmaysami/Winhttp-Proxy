# Winhttp-Proxy

### Automatic Proxy Configuration in VBA WinHTTP 5.1 (REST API)
Applications that port from WinINet to WinHTTP may need to use the same autoproxy settings that they can retrieve under WinINet or Internet Explorer (IE). The WinHTTP version 5.1 API can retrieve and use these proxy settings. In general, WinHTTP specifies the proxy and proxy bypass servers on a per-session basis when the session is created. These settings can be overridden on a per-request basis.

To use the same proxy configuration as WinINet or IE, the WinHTTP client should set proxy settings for the session. In addition, if IE or WinINet are configured to use Web Proxy Auto-Discovery (WPAD), the WinHTTP client that uses those settings must set proxy settings on a per-request basis. 

This package provides VB utilities (based on a code shared by Stephen Sulzer @ssulzer) to automatically detect Connection Proxy configuration that can be passed down to be used for Winhttp request object. It supports both 32-bits and 64-bits platforms of Windows and MS-Office and will detect Various Proxy Configurations (Auto Detect, Auto Config URL PAC, Proxy ...).

### Usage:
GetProxyInfoForUrl function, when or without optional input arguments. The return value is of type of *ProxyInfo* defined in the script.

```
Public Function GetProxyInfoForUrl(Optional URL, Optional ProxyDetails As Variant) As ProxyInfo

Public Type ProxyInfo
   ProxyActive As Boolean
   ProxyServer As String
   ProxyBypass As String
End Type
```

> Syntax1: GetProxyInfoForUrl()
>
> Syntax2: GetProxyInfoForUrl("http://www.google.com", ProxyDetails)
>
> Syntax3: GetProxyInfoForUrl(Array("http://www.google.com", "http://www.microsoft.com"), ProxyDetails)
>
> Inputs:
>
>   opt IN  URL(s)      : Array of or Single String Full URLs to AutoDetect Proxy
>
>   opt OUT ProxyDetails: Custom IE Proxy Structure to Pass out IE Proxy Details and Status Code
>
>       (1) = IE AutoDetect    (fAutoDetect)
>
>       (2) = IE AutoCofigUrl  (lpszAutoConfigUrl)
>
>       (3) = IE Proxy         (lpszProxy)
>
>       (4) = IE Proxy Bypass  (lpszProxyBypass)
>
>       (5) = DevCode



### References:
 - [Microsoft Article](https://docs.microsoft.com/en-us/windows/desktop/winhttp/setting-wininet-proxy-configurations-in-winhttp)
