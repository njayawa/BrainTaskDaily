Attribute VB_Name = "Module1"
Option Explicit

Public fMainForm As frmMain
Public fTip As New frmTip
Public fCat As New frmCategory
Public Declare Function time Lib "msvcrt" (ByRef ptr As Long) As Long
Type Tip
        Name As String
        Text As String
End Type

Sub GenerateTips(doc As frmDocument)
   
End Sub
Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash


    fMainForm.Show
End Sub

'*******************************************************************

'Purpose   : Stores date as long integer

'Receives : String implementation of date value

'Returnes  : Long integer representing provided date

'------------------------------------------------------------------------------------------
'********************************************************************

Function Date2Num(strDate As String) As Long

    On Error GoTo err_hndl

    Dim lngResult As Long



    lngResult = CDate(strDate) - 29220

    MinMaxLong lngResult, -(2 ^ 15 - 2), 2 ^ 15 - 1

    Date2Num = lngResult



    Exit Function

err_hndl:

    Date2Num = 0

End Function





'*******************************************************************

'Purpose   : Limits value inside provided ranks

'Receives : Long integers of value and its ranks

'Returnes  : Long integer

'------------------------------------------------------------------------------------------


'********************************************************************

Public Sub MinMaxLong(A&, B&, C&)

    If A& < B& Then A& = B&

    If A& > C& Then A& = C&

End Sub





'*******************************************************************

'Purpose   : Converts Long into string representing date

'Receives : Long integer implementing date value

'Returnes  : String in format dd/mm/yyyy

'------------------------------------------------------------------------------------------

'Created by Vital on 14/02/1997

'********************************************************************

Function Num2Date(lngCurDat) As String

    On Error GoTo err_hndl

    Num2Date$ = Format(lngCurDat + 29220, "dd/MM/yyyy")

    

    Exit Function

err_hndl:

    Num2Date = ""

End Function
