Attribute VB_Name = "ctrXMLUtils"
'========================================================================================
'                               CLASSter 2.0
'                    from URFIN JUS (www.urfinjus.net)
'                        XML manipulation utilities
'                  Copyright 2002. All rights reserved.
'=========================================================================================

Option Explicit

Public Function xmlSaveRecordset(ByVal RS As ADODB.Recordset) As String
    On Error GoTo errHandler
    Dim S As ADODB.Stream
    On Error GoTo errHandler
    Set S = New ADODB.Stream
    S.Open
    RS.Save S, adPersistXML
    S.Position = 0
    xmlSaveRecordset = S.ReadText
    Exit Function
errHandler:
    Err.Raise Err.Number, Err.Source, "(xmlSaveRecordset) " & Err.Description
End Function

'Extracts z:row elements, changes [z:row] to ZRowElem
Public Function xmlExtractRecords(ByVal RS As ADODB.Recordset, ByVal ZRowElem As String) As String
    On Error GoTo errHandler
    Dim P1 As Long, P2 As Long, xml As String
    xml = xmlSaveRecordset(RS)
    P1 = InStr(1, xml, "<rs:data>")
    P2 = InStr(1, xml, "</rs:data>")
    If P2 > P1 And P1 > 0 Then
        P1 = P1 + Len("<rs:data>")
        xml = Mid$(xml, P1, P2 - P1)
        xmlExtractRecords = xmlChangeElemName(xml, "z:row", ZRowElem)
        Else
        xmlExtractRecords = ""
    End If
    Exit Function
errHandler:
    Err.Raise Err.Number, Err.Source, "(xmlExtractRecords) " & Err.Description
End Function

'We can use string matching, because patterns include special symbols that
'are always escaped in XML attribute values
Public Function xmlChangeElemName(ByVal xml As String, _
                                ByVal OldName As String, _
                                ByVal NewName As String) As String
    xml = Replace(xml, "<" & OldName & " ", "<" & NewName & " ")
    xml = Replace(xml, "<" & OldName & ">", "<" & NewName & ">")
    xml = Replace(xml, "<" & OldName & "/>", "<" & NewName & "/>")
    xml = Replace(xml, "</" & OldName & ">", "</" & NewName & ">")
    xmlChangeElemName = xml
End Function

Public Function xmlMakeElement(ByVal Elem As String, ByVal Attrs As String, ByVal Content As String) As String
    xmlMakeElement = "<" & Elem & Attrs & IIf(Content = "", "/>", ">" & Content & "</" & Elem & ">") & vbNewLine
End Function

Public Function xmlMakeAttr(ByVal Attr As String, ByVal Value As String) As String
    xmlMakeAttr = " " & Attr & "='" & XMLEscape(Value) & "'"
End Function

Public Function XMLEscape(ByVal S As String) As String
    S = Replace(S, "&", "&amp;")
    S = Replace(S, "<", "&lt;")
    S = Replace(S, ">", "&gt;")
    S = Replace(S, """", "&quot;")
    S = Replace(S, "'", "&apos;")
    XMLEscape = S
End Function

Public Function xmlEnclose(ByVal xmlSrc As String, ByVal ElemOrTag As String, _
        Optional ByVal ClosingTag As String) As String
    Dim xmlRes As String
    If ElemOrTag = "" Then
        xmlRes = xmlSrc
        Else
        If InStr(1, ElemOrTag, "<") = 0 Then
            xmlRes = "<" & ElemOrTag & ">" & xmlSrc
            Else
            xmlRes = ElemOrTag & xmlSrc
        End If
        If ClosingTag = "" Then
            xmlRes = xmlRes & "</" & ElemOrTag & ">"
            Else
            xmlRes = xmlRes & ClosingTag
        End If
    End If
    xmlEnclose = xmlRes
End Function



