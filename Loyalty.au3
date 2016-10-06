#include <Array.au3>
#include <FileConstants.au3>
#include "XML.au3"
#include <GuiListView.au3>

Global Const $HTTP_STATUS_OK = 200


$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")    ; Initialize a COM error handler
; This is my custom defined error handler



#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 615, 437, 192, 124)
$Label5 = GUICtrlCreateLabel("Proxy:", 280, 110, 30, 17)
$Input1 = GUICtrlCreateInput("172.16.6.1:8080", 320, 107, 100, 21)
$Button1 = GUICtrlCreateButton("Выполнить запрос", 450, 90, 123, 49, $BS_MULTILINE)
$Group1 = GUICtrlCreateGroup("Отчетный период", 16, 72, 233, 89)
$Label1 = GUICtrlCreateLabel("Начало:", 24, 96, 44, 25)
$Date1 = GUICtrlCreateDate("", 80, 96, 154, 21, $DTS_SHORTDATEFORMAT)
$Date2 = GUICtrlCreateDate("", 80, 128, 154, 21, $DTS_SHORTDATEFORMAT)
$Label2 = GUICtrlCreateLabel("Конец:", 24, 128, 38, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Label3 = GUICtrlCreateLabel("Опрос (процедура): Пожалуйста, оцените качество приема у врача", 72, 32, 508, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("SaaS: Loyalty Reporter", 208, 8, 190, 24)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$ListView1 = GUICtrlCreateListView("POS|Сотрудник|Отлично|Хорошо|Затрудняюсь ответить|Не очень|" & _
	"Плохо|% Отлично|% Хорошо|% Затрудняюсь ответить|% Не очень|% Плохо", 16, 176, 585, 249)
Local $width1 = 100
Local $width2 = 35
_GUICtrlListView_SetColumnWidth($ListView1, 0, $width1)
_GUICtrlListView_SetColumnWidth($ListView1, 1, $width1)
_GUICtrlListView_SetColumnWidth($ListView1, 2, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 3, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 4, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 5, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 6, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 7, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 8, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 9, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 10, $width2)
_GUICtrlListView_SetColumnWidth($ListView1, 11, $width2)
GUICtrlSetState($Button1, $GUI_FOCUS)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1
			Local $strDate1 = GUICtrlRead($Date1)
			Local $strDate2 = GUICtrlRead($Date2)
			Local $strProxy = GUICtrlRead($Input1)
			ConsoleWrite($strDate1 & " " & $strDate2 & @CRLF)
			ButtonPressed($strDate1, $strDate2, $strProxy)
	EndSwitch
WEnd



Func ButtonPressed($strDate1, $strDate2, $strProxy)
	Local $strUrl = ""
	Local $strReportID = ""
	Local $strDtBegin = $strDate1 & " 00:00:00"
	Local $strDtEnd = $strDate2 & " 23:59:59"
	Local $strUseUTC = "0"
	Local $strLoginName = ""
	Local $strPassword = ""
	Local $nQuestionID =

	Local $strBody = "LoginName=" & $strLoginName & _
					 "&Password=" & $strPassword & _
					 "&QuestionID=" & $nQuestionID & _
					 "&ReportID=" & $strReportID & _
					 "&Begin=" & $strDtBegin & _
					 "&End=" & $strDtEnd & _
					 "&UseUTC=" & $strUseUTC

;~ 	MsgBox(0, "", $strBody)

	Local $strXmlResponse = HttpPost($strUrl, $strBody, $strProxy)

	Local $strFileName = @ScriptDir & "\response.xml"
	Local $hFile = FileOpen($strFileName, BitOR($FO_OVERWRITE, $FO_ANSI))
	FileWrite($hFile, $strXmlResponse)
	FileClose($hFile)

	Local $resultArray = Example_1__XML_SelectNodes($strFileName)
	;~ _ArrayDisplay($resultArray)

	Local $result[0][12]
	If IsArray($resultArray) Then
;~ 		ConsoleWrite("Array" & @CRLF)
		For $i = 1 To UBound($resultArray, $UBOUND_ROWS) - 1
			Local $tmpArray[1][12]
			For $x = 0 To 11
				$tmpArray[0][$x] = $resultArray[$i + $x][3]
			Next
			_ArrayAdd($result, $tmpArray)
			$i += 11
		Next
	Else
;~ 		ConsoleWrite("not Array" & @CRLF)
		Local $tmp[1][12]
		$tmp[0][0] = "Нет данных"
		_ArrayAdd($result, $tmp)
	EndIf

;~ 	_ArrayDisplay($result)


	_GUICtrlListView_DeleteAllItems($ListView1)
	_GUICtrlListView_AddArray($ListView1, $result)
EndFunc




Func Example_1__XML_SelectNodes($strFileName)

	; first you must create $oXmlDoc object
	Local $oXMLDoc = _XML_CreateDOMDocument(Default)
	If @error Then
		Return
;~ 		MsgBox(0, '_XML_CreateDOMDocument @error:', XML_My_ErrorParser(@error))
	Else
		; now you can add EVENT Handler
		Local $oXMLDOM_EventsHandler = ObjEvent($oXMLDoc, "XML_DOM_EVENT_")
		#forceref $oXMLDOM_EventsHandler

		; Load file to $oXmlDoc
		_XML_Load($oXMLDoc, $strFileName)
		If @error Then
			Return
;~ 			MsgBox(0, '_XML_Load @error:', XML_My_ErrorParser(@error))
		Else

			; simple display $oXmlDoc - for checking only
			Local $sXmlAfterTidy = _XML_TIDY($oXMLDoc)
			If @error Then
				Return
;~ 				MsgBox(0, '_XML_TIDY @error:', XML_My_ErrorParser(@error))
;~ 			Else
;~ 				MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, 'Example_1__XML_SelectNodes', $sXmlAfterTidy)
			EndIf

			; selecting nodes
			Local $oNodesColl = _XML_SelectNodes($oXMLDoc, "//Rows/Row/Cell")
			If @error Then
				Return
;~ 				MsgBox(0, '_XML_SelectNodes @error:', XML_My_ErrorParser(@error))
			Else
				; change Nodes Collection to array
				Local $aNodesColl = _XML_Array_GetNodesProperties($oNodesColl)
				If @error Then
					Return
;~ 					MsgBox(0, '_XML_Array_GetNodesProperties @error:', XML_My_ErrorParser(@error))
				Else
					Return($aNodesColl)
				EndIf
			EndIf
		EndIf
	EndIf
EndFunc    ;==>Example_1__XML_SelectNodes


; #FUNCTION# ====================================================================================================================
; Name ..........: XML_My_ErrorParser
; Description ...: Changing $XML_ERR_ ... to human readable description
; Syntax ........: XML_My_ErrorParser($iXMLWrapper_Error, $iXMLWrapper_Extended)
; Parameters ....: $iXMLWrapper_Error	- an integer value.
;                  $iXMLWrapper_Extended           - an integer value.
; Return values .: description as string
; Author ........: mLipok
; Modified ......:
; Remarks .......: This function is only example of how user can parse @error and @extended to human readable description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func XML_My_ErrorParser($iXMLWrapper_Error, $iXMLWrapper_Extended = 0)
	Local $sErrorInfo = ''
	Switch $iXMLWrapper_Error
		Case $XML_ERR_OK
			$sErrorInfo = '$XML_ERR_OK=' & $XML_ERR_OK & @CRLF & 'All is ok.'
		Case $XML_ERR_GENERAL
			$sErrorInfo = '$XML_ERR_GENERAL=' & $XML_ERR_GENERAL & @CRLF & 'The error which is not specifically defined.'
		Case $XML_ERR_COMERROR
			$sErrorInfo = '$XML_ERR_COMERROR=' & $XML_ERR_COMERROR & @CRLF & 'COM ERROR OCCURED. Check @extended and your own error handler function for details.'
		Case $XML_ERR_ISNOTOBJECT
			$sErrorInfo = '$XML_ERR_ISNOTOBJECT=' & $XML_ERR_ISNOTOBJECT & @CRLF & 'No object passed to function'
		Case $XML_ERR_INVALIDDOMDOC
			$sErrorInfo = '$XML_ERR_INVALIDDOMDOC=' & $XML_ERR_INVALIDDOMDOC & @CRLF & 'Invalid object passed to function'
		Case $XML_ERR_INVALIDATTRIB
			$sErrorInfo = '$XML_ERR_INVALIDATTRIB=' & $XML_ERR_INVALIDATTRIB & @CRLF & 'Invalid object passed to function.'
		Case $XML_ERR_INVALIDNODETYPE
			$sErrorInfo = '$XML_ERR_INVALIDNODETYPE=' & $XML_ERR_INVALIDNODETYPE & @CRLF & 'Invalid object passed to function.'
		Case $XML_ERR_OBJCREATE
			$sErrorInfo = '$XML_ERR_OBJCREATE=' & $XML_ERR_OBJCREATE & @CRLF & 'Object can not be created.'
		Case $XML_ERR_NODECREATE
			$sErrorInfo = '$XML_ERR_NODECREATE=' & $XML_ERR_NODECREATE & @CRLF & 'Can not create Node - check also COM Error Handler'
		Case $XML_ERR_NODEAPPEND
			$sErrorInfo = '$XML_ERR_NODEAPPEND=' & $XML_ERR_NODEAPPEND & @CRLF & 'Can not append Node - check also COM Error Handler'
		Case $XML_ERR_PARSE
			$sErrorInfo = '$XML_ERR_PARSE=' & $XML_ERR_PARSE & @CRLF & 'Error: with Parsing objects, .parseError.errorCode=' & $iXMLWrapper_Extended & ' Use _XML_ErrorParser_GetDescription() for get details.'
		Case $XML_ERR_PARSE_XSL
			$sErrorInfo = '$XML_ERR_PARSE_XSL=' & $XML_ERR_PARSE_XSL & @CRLF & 'Error with Parsing XSL objects .parseError.errorCode=' & $iXMLWrapper_Extended & ' Use _XML_ErrorParser_GetDescription() for get details.'
		Case $XML_ERR_LOAD
			$sErrorInfo = '$XML_ERR_LOAD=' & $XML_ERR_LOAD & @CRLF & 'Error opening specified file.'
		Case $XML_ERR_SAVE
			$sErrorInfo = '$XML_ERR_SAVE=' & $XML_ERR_SAVE & @CRLF & 'Error saving file.'
		Case $XML_ERR_PARAMETER
			$sErrorInfo = '$XML_ERR_PARAMETER=' & $XML_ERR_PARAMETER & @CRLF & 'Wrong parameter passed to function.'
		Case $XML_ERR_ARRAY
			$sErrorInfo = '$XML_ERR_ARRAY=' & $XML_ERR_ARRAY & @CRLF & 'Wrong array parameter passed to function. Check array dimension and conent.'
		Case $XML_ERR_XPATH
			$sErrorInfo = '$XML_ERR_XPATH=' & $XML_ERR_XPATH & @CRLF & 'XPath syntax error - check also COM Error Handler.'
		Case $XML_ERR_NONODESMATCH
			$sErrorInfo = '$XML_ERR_NONODESMATCH=' & $XML_ERR_NONODESMATCH & @CRLF & 'No nodes match the XPath expression'
		Case $XML_ERR_NOCHILDMATCH
			$sErrorInfo = '$XML_ERR_NOCHILDMATCH=' & $XML_ERR_NOCHILDMATCH & @CRLF & 'There is no Child in nodes matched by XPath expression.'
		Case $XML_ERR_NOATTRMATCH
			$sErrorInfo = '$XML_ERR_NOATTRMATCH=' & $XML_ERR_NOATTRMATCH & @CRLF & 'There is no such attribute in selected node.'
		Case $XML_ERR_DOMVERSION
			$sErrorInfo = '$XML_ERR_DOMVERSION=' & $XML_ERR_DOMVERSION & @CRLF & 'DOM Version: ' & 'MSXML Version ' & $iXMLWrapper_Extended & ' or greater required for this function'
		Case $XML_ERR_EMPTYCOLLECTION
			$sErrorInfo = '$XML_ERR_EMPTYCOLLECTION=' & $XML_ERR_EMPTYCOLLECTION & @CRLF & 'Collections of objects was empty'
		Case $XML_ERR_EMPTYOBJECT
			$sErrorInfo = '$XML_ERR_EMPTYOBJECT=' & $XML_ERR_EMPTYOBJECT & @CRLF & 'Object is empty'
		Case Else
			$sErrorInfo = '=' & $iXMLWrapper_Error & @CRLF & 'NO ERROR DESCRIPTION FOR THIS @error'
	EndSwitch

	Local $sExtendedInfo = ''
	Switch $iXMLWrapper_Error
		Case $XML_ERR_COMERROR, $XML_ERR_NODEAPPEND, $XML_ERR_NODECREATE
			$sExtendedInfo = 'COM ERROR NUMBER (@error returned via @extended) =' & $iXMLWrapper_Extended
		Case $XML_ERR_PARAMETER
			$sExtendedInfo = 'This @error was fired by parameter: #' & $iXMLWrapper_Extended
		Case Else
			Switch $iXMLWrapper_Extended
				Case $XML_EXT_DEFAULT
					$sExtendedInfo = '$XML_EXT_DEFAULT=' & $XML_EXT_DEFAULT & @CRLF & 'Default - Do not return any additional information'
				Case $XML_EXT_XMLDOM
					$sExtendedInfo = '$XML_EXT_XMLDOM=' & $XML_EXT_XMLDOM & @CRLF & '"Microsoft.XMLDOM" related Error'
				Case $XML_EXT_DOMDOCUMENT
					$sExtendedInfo = '$XML_EXT_DOMDOCUMENT=' & $XML_EXT_DOMDOCUMENT & @CRLF & '"Msxml2.DOMDocument" related Error'
				Case $XML_EXT_XSLTEMPLATE
					$sExtendedInfo = '$XML_EXT_XSLTEMPLATE=' & $XML_EXT_XSLTEMPLATE & @CRLF & '"Msxml2.XSLTemplate" related Error'
				Case $XML_EXT_SAXXMLREADER
					$sExtendedInfo = '$XML_EXT_SAXXMLREADER=' & $XML_EXT_SAXXMLREADER & @CRLF & '"MSXML2.SAXXMLReader" related Error'
				Case $XML_EXT_MXXMLWRITER
					$sExtendedInfo = '$XML_EXT_MXXMLWRITER=' & $XML_EXT_MXXMLWRITER & @CRLF & '"MSXML2.MXXMLWriter" related Error'
				Case $XML_EXT_FREETHREADEDDOMDOCUMENT
					$sExtendedInfo = '$XML_EXT_FREETHREADEDDOMDOCUMENT=' & $XML_EXT_FREETHREADEDDOMDOCUMENT & @CRLF & '"Msxml2.FreeThreadedDOMDocument" related Error'
				Case $XML_EXT_XMLSCHEMACACHE
					$sExtendedInfo = '$XML_EXT_XMLSCHEMACACHE=' & $XML_EXT_XMLSCHEMACACHE & @CRLF & '"Msxml2.XMLSchemaCache." related Error'
				Case $XML_EXT_STREAM
					$sExtendedInfo = '$XML_EXT_STREAM=' & $XML_EXT_STREAM & @CRLF & '"ADODB.STREAM" related Error'
				Case $XML_EXT_ENCODING
					$sExtendedInfo = '$XML_EXT_ENCODING=' & $XML_EXT_ENCODING & @CRLF & 'Encoding related Error'
				Case Else
					$sExtendedInfo = '$iXMLWrapper_Extended=' & $iXMLWrapper_Extended & @CRLF & 'NO ERROR DESCRIPTION FOR THIS @extened'
			EndSwitch
	EndSwitch
	; return back @error and @extended for further debuging
	Return SetError($iXMLWrapper_Error, $iXMLWrapper_Extended, _
			'@error description:' & @CRLF & _
			$sErrorInfo & @CRLF & _
			@CRLF & _
			'@extended description:' & @CRLF & _
			$sExtendedInfo & @CRLF & _
			'')

EndFunc    ;==>XML_My_ErrorParser

Func ErrFunc_CustomUserHandler_XML($oError)

	; here is declared another path to UDF au3 file
	; thanks to this with using _XML_ComErrorHandler_UserFunction(ErrFunc_CustomUserHandler_XML)
	;  you get errors which after pressing F4 in SciTE4AutoIt you goes directly to the specified UDF Error Line
	ConsoleWrite(@ScriptDir & '\XMLWrapperEx.au3' & " (" & $oError.scriptline & ") : UDF ==> COM Error intercepted ! " & @CRLF & _
			@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
			@TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
			@TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
			@TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
			@TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
			@TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
			@TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
			@TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
			@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc    ;==>ErrFunc_CustomUserHandler_XML


Func HttpPost($sURL, $sData = "", $strProxy = "")
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
	ConsoleWrite("obj create" & @CRLF)

	If $strProxy Then
		$oHTTP.SetProxy(2, $strProxy)
		ConsoleWrite("obj proxy" & @CRLF)
	EndIf

	$oHTTP.Open("POST", $sURL, False)
	If (@error) Then Return SetError(1, 0, 0)

	ConsoleWrite("obj open" & @CRLF)

	$oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	$oHTTP.SetRequestHeader("RequestType", "GetXmlLoyaltyReport")
	$oHTTP.SetRequestHeader("Content-Length", StringLen($sData))

	ConsoleWrite("obj header" & @CRLF)

	$oHTTP.Send($sData)
	If (@error) Then Return SetError(2, 0, 0)

	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc


Func MyErrFunc()
  Msgbox(0,"AutoItCOM Test","We intercepted a COM Error !"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
Endfunc