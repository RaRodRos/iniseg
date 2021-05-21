' Attribute VB_Name = "RaUI"
Option Explicit

' En este módulo se implementarán todas las subrutinas y funciones a las que
	' pueda interesar acceder desde la UI

Sub uiFileSaveAsNew()
    RaMacros.FileSaveAsNew _
        dcArg:=ActiveDocument, _
        rgArg:=Nothing, _
        stNewName:=vbNullString, _
        stPrefix:=vbNullString, _
        stSuffix:=vbNullString, _
        stPath:=vbNullString, _
        bClose:=True, _
        bCompatibility:=False, _
        bVisible:=False
End Sub

Sub uiFileCopy()
    RaMacros.FileCopy _
        dcArg:=ActiveDocument, _
        stPrefix:=vbNullString, _
        stSuffix:=vbNullString, _
        stPath:=vbNullString
End Sub