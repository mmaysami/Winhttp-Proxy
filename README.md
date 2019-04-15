# Winhttp-Proxy

## Setting Automatic Proxy on WinHTTP 5.1 (VB-REST)
Applications that port from WinINet to WinHTTP may need to use the same autoproxy settings that they can retrieve under WinINet or Internet Explorer (IE). The WinHTTP version 5.1 API can retrieve and use these proxy settings. In general, WinHTTP specifies the proxy and proxy bypass servers on a per-session basis when the session is created. These settings can be overridden on a per-request basis.

To use the same proxy configuration as WinINet or IE, the WinHTTP client should set proxy settings for the session. In addition, if IE or WinINet are configured to use Web Proxy Auto-Discovery (WPAD), the WinHTTP client that uses those settings must set proxy settings on a per-request basis. 

This package provides VB utilities (based on a code shared by Stephen Sulzer @ssulzer) to automatically detect Connection Proxy configuration that can be passed down to be used for Winhttp request object. It supports both 32-bits and 64-bits platforms of Windows and MS-Office and will detect Various Proxy Configurations (Auto Detect, Auto Config URL PAC, Proxy ...).

### References:
 - [Microsoft Article](https://docs.microsoft.com/en-us/windows/desktop/winhttp/setting-wininet-proxy-configurations-in-winhttp)
