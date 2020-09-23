VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CUri
' Author: Jon Davis (from scratch)

Private sURI As String
Private sProtocol As String
Private sHost As String
Private sUser As String
Private sPwd As String
Private sRes As String
Private iPort As Integer
Private sResQuery As String
Private bInit As Boolean

Private Const CURI_DEFAULT_TCPPROTOCOL = "http"

Private Function getPortFromProtocol(sProtocol As String) As Integer
        Select Case sProtocol
        Case "http"
            getPortFromProtocol = 80
        Case "https"
            getPortFromProtocol = 443
        Case "ftp"
            getPortFromProtocol = 21
        Case "telnet"
            getPortFromProtocol = 23
        Case "finger"
            getPortFromProtocol = 79
        Case "irc"
            getPortFromProtocol = 6667
        'Case etc...
        End Select
End Function

Public Property Get Initialized() As Boolean
    Initialized = bInit
End Property

Public Property Get Uri() As String
Attribute Uri.VB_UserMemId = 0
    Uri = sURI
End Property

Public Property Let Uri(sNewUri As String)
    If sNewUri = "" Then Exit Property
    sURI = sNewUri
    sProtocol = parseProtocol(sNewUri)
    sHost = parseHost(sNewUri)
    sUser = parseUsername(sNewUri)
    sPwd = parsePassword(sNewUri)
    sRes = parseResource(sNewUri)
    iPort = parsePort(sNewUri)
    sResQuery = parseQuery(sNewUri)
    bInit = True
End Property

Public Property Get Protocol() As String
    Protocol = sProtocol
End Property

Public Property Get Host() As String
    Host = sHost
End Property

Public Property Get Resource() As String
    Resource = sRes
End Property

Public Property Get FullResource() As String
    If sResQuery <> "" Then
        FullResource = sRes & "?" & sResQuery
    Else
        FullResource = sRes
    End If
End Property

Public Property Get QueryString() As String
    QueryString = sResQuery
End Property

Public Property Get Username() As String
    Username = sUser
End Property

Public Property Get Password() As String
    Password = sPwd
End Property

Public Property Get Port() As Integer
    Port = iPort
End Property

Private Function parseProtocol(sURI As String)
    Dim slashLoc As Integer
    Dim colonLoc As Integer
    Dim protLoc As Integer
    sURI = Replace(sURI, "\", "/")
    
    'any slashes?
    slashLoc = InStr(sURI, "/")
    colonLoc = InStr(sURI, ":")
    If (slashLoc = 0 And colonLoc = 0) Or _
     (slashLoc > 0 And slashLoc <= colonLoc + 1) Then
        'first slashes at protocol delim?
        protLoc = InStr(sURI, "://")
        If protLoc > 0 And protLoc <= slashLoc - 1 Then
            parseProtocol = Left(sURI, protLoc - 1)
        Else
            parseProtocol = CURI_DEFAULT_TCPPROTOCOL
        End If
    Else ' no slashes at colon
        If InStr(Left(sURI, colonLoc), ".") > 0 Or _
           InStr(Left(sURI, colonLoc), "@") > 0 Then
            ' There is a dot-delimited hostname or a
            ' username/password combination before
            ' the colon, so ignore colon because
            ' no protocol was specified.
            ' Watching for dots is not URI conformant but
            ' neither is leaving out the protocol in the first place
            colonLoc = 0
        End If
        protLoc = colonLoc
        If protLoc > 0 Then
            parseProtocol = Left(sURI, protLoc - 1)
        Else
            parseProtocol = CURI_DEFAULT_TCPPROTOCOL
        End If
    End If
End Function

Private Function parseHost(sURI As String) As String
    Dim hostUrl As String
    hostUrl = parseHostUrl(sURI)
    If InStr(hostUrl, "@") > 0 Then
        parseHost = Right(hostUrl, Len(hostUrl) - InStr(hostUrl, "@"))
    Else
        parseHost = hostUrl
    End If
    If InStr(parseHost, ":") > 0 Then
        'there's a port number in there!
        parseHost = Left(parseHost, InStr(parseHost, ":") - 1)
    End If
End Function

Private Function parseHostUrl(sURI As String)
    Dim sProt As String
    Dim iProtIdentLen As Integer
    Dim iProtLen As Integer
    Dim endPoint As Integer
    sURI = Replace(sURI, "\", "/")
    sProt = parseProtocol(sURI)
    'sURI is in the format [protocol][?][?]........
    If InStr(Len(sProt), sURI, "://") = Len(sProt) + 1 Then
        'sURI is in the format [protocol]://[?]........
        iProtIdentLen = 3
    ElseIf InStr(Len(sProt), sURI, ":") = Len(sProt) + 1 Then
        'sURI is in the format [protocol]:[?]........
        iProtIdentLen = 1
    Else
        'sURI is in the format [?] (no protocol specified)
        iProtIdentLen = Len(sProt) * -1
    End If
    iProtLen = Len(sProt) + iProtIdentLen
    endPoint = InStr(iProtLen + 1, sURI, "/")
    If endPoint <= 0 Then endPoint = Len(sURI) + 1
    parseHostUrl = Mid(sURI, iProtLen + 1, endPoint - iProtLen - 1)
    'If IsNumeric(parseHostUrl) Then 'oops, we got the port
    'End If
End Function

Private Function parsePort(sURI As String) As Integer
    Dim hostUrl As String
    hostUrl = parseHostUrl(sURI)
    If InStr(hostUrl, "@") > 0 Then
        hostUrl = Right(hostUrl, Len(hostUrl) - InStr(hostUrl, "@"))
    Else
        hostUrl = hostUrl
    End If
    If InStr(hostUrl, ":") > 0 Then
        'there's a port number in there!
        hostUrl = Right(hostUrl, Len(hostUrl) - InStr(hostUrl, ":"))
        If hostUrl <> "" And IsNumeric(hostUrl) Then
            parsePort = CInt(hostUrl)
        End If
    End If
    If parsePort = 0 Then
        parsePort = getPortFromProtocol(parseProtocol(sURI))
    End If
End Function

Private Function parseResource(sURI As String) As String
    Dim ResLoc As Integer
    Dim hu As String
    sURI = Replace(sURI, "\", "/")
    hu = parseHostUrl(sURI)
    ResLoc = InStr(sURI, hu) + Len(hu) - 1
    If ResLoc < Len(sURI) Then
        parseResource = Right(sURI, Len(sURI) - ResLoc)
    Else
        parseResource = "/"
    End If
    
    'extract query
    Dim queryLoc As Integer
    queryLoc = InStr(parseResource, "?")
    If queryLoc > 0 Then
        parseResource = Left(parseResource, queryLoc - 1)
    End If
End Function

Private Function parseQuery(sURI As String) As String
    Dim ResLoc As Integer
    Dim hu As String
    sURI = Replace(sURI, "\", "/")
    hu = parseHostUrl(sURI)
    ResLoc = InStr(sURI, hu) + Len(hu) - 1
    If ResLoc < Len(sURI) Then
        parseQuery = Right(sURI, Len(sURI) - ResLoc)
    Else
        parseQuery = "/"
    End If
    
    'extract query
    Dim queryLoc As Integer
    queryLoc = InStr(parseQuery, "?")
    If queryLoc > 0 Then
        parseQuery = Right(parseQuery, Len(parseQuery) - queryLoc)
    Else
        parseQuery = ""
    End If
End Function

Private Function parseUsername(sURI As String) As String
    Dim u_p As String
    Dim hu As String
    Dim huSep As Integer
    Dim upSep As Integer
    
    hu = parseHostUrl(sURI)
    huSep = InStr(hu, "@")
    If huSep > 0 Then
        u_p = Left(hu, huSep - 1)
        upSep = InStr(u_p, ":")
        If upSep > 0 Then
            parseUsername = Left(u_p, upSep - 1)
        Else
            parseUsername = u_p
        End If
    End If
End Function

Private Function parsePassword(sURI As String) As String
    Dim u_p As String
    Dim hu As String
    Dim huSep As Integer
    Dim upSep As Integer
    
    hu = parseHostUrl(sURI)
    huSep = InStr(hu, "@")
    If huSep > 0 Then
        u_p = Left(hu, huSep - 1)
        upSep = InStr(u_p, ":")
        If upSep > 0 Then
            parsePassword = Right(u_p, Len(u_p) - upSep)
        End If
    End If
End Function


