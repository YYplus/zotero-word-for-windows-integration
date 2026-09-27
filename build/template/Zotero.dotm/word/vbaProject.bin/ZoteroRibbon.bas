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

Private Const ZoteroRibbonLocalizationNamespace As String = "http://www.zotero.org/ribbon-localization/v1"
Private Const ZoteroRibbonDefaultLocale As String = "en-US"

Private ZoteroRibbonLocalizationPart As CustomXMLPart
Private ZoteroRibbonLocale As String

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

Private Function ZoteroRibbonGetLocalizationPart() As CustomXMLPart
    Dim matchingParts As CustomXMLParts

    If ZoteroRibbonLocalizationPart Is Nothing Then
        Set matchingParts = ThisDocument.CustomXMLParts.SelectByNamespace( _
            ZoteroRibbonLocalizationNamespace)
        If matchingParts.Count > 0 Then
            Set ZoteroRibbonLocalizationPart = matchingParts(1)
        End If
    End If

    If Not ZoteroRibbonLocalizationPart Is Nothing Then
        Set ZoteroRibbonGetLocalizationPart = ZoteroRibbonLocalizationPart
    End If
End Function

Private Function ZoteroRibbonGetDefaultLocale() As String
    Dim localizationPart As CustomXMLPart
    Dim localeNode As CustomXMLNode

    Set localizationPart = ZoteroRibbonGetLocalizationPart()
    If localizationPart Is Nothing Then
        ZoteroRibbonGetDefaultLocale = ZoteroRibbonDefaultLocale
        Exit Function
    End If

    Set localeNode = localizationPart.SelectSingleNode( _
        "/*[local-name()='ribbonLocalization']/@defaultLocale")
    If localeNode Is Nothing Then
        ZoteroRibbonGetDefaultLocale = ZoteroRibbonDefaultLocale
    Else
        ZoteroRibbonGetDefaultLocale = localeNode.Text
    End If
End Function

Private Function ZoteroRibbonResolveLocale() As String
    Dim localizationPart As CustomXMLPart
    Dim localeNode As CustomXMLNode
    Dim languageID As Long

    If Len(ZoteroRibbonLocale) > 0 Then
        ZoteroRibbonResolveLocale = ZoteroRibbonLocale
        Exit Function
    End If

    Set localizationPart = ZoteroRibbonGetLocalizationPart()
    If localizationPart Is Nothing Then
        ZoteroRibbonLocale = ZoteroRibbonDefaultLocale
        ZoteroRibbonResolveLocale = ZoteroRibbonLocale
        Exit Function
    End If

    On Error Resume Next
    languageID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    On Error GoTo 0

    If languageID <> 0 Then
        Set localeNode = localizationPart.SelectSingleNode( _
            "/*[local-name()='ribbonLocalization']" & _
            "/*[local-name()='languages']" & _
            "/*[local-name()='language' and @officeLanguageID='" & _
            CStr(languageID) & "']/@locale")
    End If

    If localeNode Is Nothing Then
        ZoteroRibbonLocale = ZoteroRibbonGetDefaultLocale()
    Else
        ZoteroRibbonLocale = localeNode.Text
    End If

    ZoteroRibbonResolveLocale = ZoteroRibbonLocale
End Function

Private Function ZoteroRibbonLookupText(controlID As String, textType As String) As String
    Dim localizationPart As CustomXMLPart
    Dim textNode As CustomXMLNode
    Dim locale As String
    Dim defaultLocale As String
    Dim xpath As String

    On Error GoTo LookupFailed

    Set localizationPart = ZoteroRibbonGetLocalizationPart()
    If localizationPart Is Nothing Then Exit Function

    locale = ZoteroRibbonResolveLocale()
    xpath = "/*[local-name()='ribbonLocalization']" & _
        "/*[local-name()='locales']" & _
        "/*[local-name()='locale' and @id='" & locale & "']" & _
        "/*[local-name()='control' and @id='" & controlID & "']" & _
        "/*[local-name()='" & textType & "']"
    Set textNode = localizationPart.SelectSingleNode(xpath)

    If textNode Is Nothing Then
        defaultLocale = ZoteroRibbonGetDefaultLocale()
        If locale <> defaultLocale Then
            xpath = "/*[local-name()='ribbonLocalization']" & _
                "/*[local-name()='locales']" & _
                "/*[local-name()='locale' and @id='" & defaultLocale & "']" & _
                "/*[local-name()='control' and @id='" & controlID & "']" & _
                "/*[local-name()='" & textType & "']"
            Set textNode = localizationPart.SelectSingleNode(xpath)
        End If
    End If

    If Not textNode Is Nothing Then
        ZoteroRibbonLookupText = textNode.Text
    End If
    Exit Function

LookupFailed:
    ZoteroRibbonLookupText = ""
End Function

Private Function ZoteroRibbonFallbackText(control As IRibbonControl, partIndex As Long) As String
    Dim parts() As String

    If Len(control.Tag) = 0 Then Exit Function
    parts = Split(control.Tag, "||")
    If partIndex >= LBound(parts) And partIndex <= UBound(parts) Then
        ZoteroRibbonFallbackText = parts(partIndex)
    End If
End Function

Sub ZoteroRibbonGetLabel(control As IRibbonControl, ByRef returnedVal)
    Dim localizedText As String

    returnedVal = ""
    On Error Resume Next
    returnedVal = ZoteroRibbonFallbackText(control, 0)
    On Error GoTo LocalizationFailed

    localizedText = ZoteroRibbonLookupText(control.Id, "label")
    If Len(localizedText) > 0 Then
        returnedVal = localizedText
    End If

LocalizationFailed:
End Sub

Sub ZoteroRibbonGetSupertip(control As IRibbonControl, ByRef returnedVal)
    Dim localizedText As String

    returnedVal = ""
    On Error Resume Next
    returnedVal = ZoteroRibbonFallbackText(control, 1)
    On Error GoTo LocalizationFailed

    localizedText = ZoteroRibbonLookupText(control.Id, "supertip")
    If Len(localizedText) > 0 Then
        returnedVal = localizedText
    End If

LocalizationFailed:
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
