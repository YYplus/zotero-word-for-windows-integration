Attribute VB_Name = "ZoteroRibbon"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Option Explicit

Sub ZoteroRibbonAddEditCitation(button As IRibbonControl)
    Call ZoteroAddEditCitation
End Sub

Sub ZoteroRibbonAddNote(button As IRibbonControl)
    Call ZoteroAddNote
End Sub

Sub ZoteroRibbonAddAnnotation(button As IRibbonControl)
    Call ZoteroAddAnnotation
End Sub

Sub ZoteroRibbonAddEditBibliography(button As IRibbonControl)
    Call ZoteroAddEditBibliography
End Sub

Sub ZoteroRibbonSetDocPrefs(button As IRibbonControl)
    Call ZoteroSetDocPrefs
End Sub

Sub ZoteroRibbonRefresh(button As IRibbonControl)
    Call ZoteroRefresh
End Sub

Sub ZoteroRibbonRemoveCodes(button As IRibbonControl)
    Call ZoteroRemoveCodes
End Sub

Private Function ZoteroRibbonUsesSimplifiedChinese() As Boolean
    Dim languageID As Long
    On Error Resume Next
    ' 2 = msoLanguageIDUI. Use the numeric value to avoid adding a dependency
    ' on a particular Office type-library version.
    languageID = Application.LanguageSettings.LanguageID(2)
    On Error GoTo 0

    ' 2052 = zh-CN; 4100 = zh-SG
    ZoteroRibbonUsesSimplifiedChinese = (languageID = 2052 Or languageID = 4100)
End Function

Private Function ZoteroRibbonTagPart(control As IRibbonControl, partIndex As Long) As String
    Dim parts() As String
    If Len(control.Tag) = 0 Then Exit Function

    ' Tag format:
    ' English label || Simplified Chinese label || English supertip || Simplified Chinese supertip
    parts = Split(control.Tag, "||")
    If partIndex >= LBound(parts) And partIndex <= UBound(parts) Then
        ZoteroRibbonTagPart = parts(partIndex)
    End If
End Function

Sub ZoteroRibbonGetLabel(control As IRibbonControl, ByRef returnedVal)
    If ZoteroRibbonUsesSimplifiedChinese() Then
        returnedVal = ZoteroRibbonTagPart(control, 1)
        If Len(returnedVal) > 0 Then Exit Sub
    End If

    returnedVal = ZoteroRibbonTagPart(control, 0)
End Sub

Sub ZoteroRibbonGetSupertip(control As IRibbonControl, ByRef returnedVal)
    If ZoteroRibbonUsesSimplifiedChinese() Then
        returnedVal = ZoteroRibbonTagPart(control, 3)
        If Len(returnedVal) > 0 Then Exit Sub
    End If

    returnedVal = ZoteroRibbonTagPart(control, 2)
End Sub

Sub ZoteroTabLabel(tb As IRibbonControl, ByRef returnedVal)
    Dim ver As Double
    ver = Val(Application.Version)
    If ver >= 15 And ver < 16 Then
        returnedVal = "ZOTERO"
    Else
        returnedVal = "Zotero"
    End If
End Sub
