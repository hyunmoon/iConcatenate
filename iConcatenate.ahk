#SingleInstance Force
#Persistent
#NoEnv
#NoTrayIcon
SetBatchLines, -1
AutoTrim Off
SetTitleMatchMode, 3
; ===========================================================================================
; This script is a GUI wrapper for ImageMagick software. It allows users to quickly concatenate image
; files which is one of many functions provided by ImageMagick.
; ImageMagick is a great software to process images but I thought its command-line interface
; is a bit difficult to use so I wrote this script to make it easier to use.
; ===========================================================================================

; Install convert.exe if it doesn't exist already
IfNotExist, convert.exe
{
	FileInstall, convert.exe, convert.exe
	if (ErrorLevel == 1) {
	MSgBox, 0, Error,
(
Failed to install [ convert.exe ]
(probably blocked by anti-virus software)

[ convert.exe ] is a harmless executable that is`na command-line tool for ImageMagick software
and it is freely available on www.imagemagick.org
)
		ExitApp
	}
}

; -------------------------------------------------------------------------------------------
; Create Gui
; -------------------------------------------------------------------------------------------
Color_Background := 0xFFFFFF
Color_Text := 0xC0C0C0
quoteMark = "

Gui Add, Radio, Section vRA1 Checked, (¡é)  Vertical
Gui Add, Radio, vRA2, (¡æ) Horizontal
GroupBox("GBRadio", "", 10, 10, "RA1|RA2")
Gui Add, Button, Section xS+119 yS+5 vSelectFiles gSelectFIlesClicked w70 h25, Select
Gui Add, Button, vConcatenate gConcatenateClicked w70 h26, Save
Gui, Font, %Color_Text%
Gui, Color, %Color_Background%
Gui, Show, AutoSize, iConcatenate
Return

; -------------------------------------------------------------------------------------------
; Select button clicked
; -------------------------------------------------------------------------------------------
SelectFIlesClicked:
{
	FileSelectFile, SelectedFiles, M3, , Choose images to concatenate, Image File(*.png; *.jpg;, *.bmp)

	If (SelectedFiles != "")
	{
		FileNames := ""
		Loop, parse, SelectedFiles, `n
		{
			if (A_Index == 1)
			{
				FileDirectory := quoteMark
				FileDirectory .= A_LoopField
			}
			else
			{
				FileNames .= FileDirectory
				FileNames .= "\"
				FileNames .= A_LoopField
				FileNames .= quoteMark
				FileNames .= A_Space
			}
		}
	}
	
	Return
}

; -------------------------------------------------------------------------------------------
; Images have been drag and dropped into Gui
; -------------------------------------------------------------------------------------------
GuiDropFiles:
{
	If (A_GuiEvent != "")
	{
		FileNames := ""
		FileList := A_GuiEvent
		Sort, FileList ; sort in alphabetical order
		
		Loop, parse, FileList, `n
		{
			FileNames .= quoteMark
			FileNames .= A_LoopField
			FileNames .= quoteMark
			FileNames .= A_Space
				
			SplitPath, A_LoopField, fileName
		}
		
		; For drag and drops, try to create image right away
		goSub, ConcatenateClicked
	}
	return
}

; -------------------------------------------------------------------------------------------
; Save button clicked
; -------------------------------------------------------------------------------------------
ConcatenateClicked:
{
	if (FileNames = "")
	{
		SoundPlay *48
		MsgBox, 0, Warning, Select images to concatenate!
		Return
	}
	
	FileSelectFile, saveFilePath, S, ,Save, Image (*.png)
	if (saveFilePath = "")
	{
		Return
	}
	else
	{
		IfNotInString, saveFilePath, .
		{
			saveFilePath .= ".png"
		}
		
		IfExist, %saveFilePath%
		{
			SplitPath, saveFilePath, saveFileName ; take file name from the full path
			
			SoundPlay *48 ; ask before overwriting
			MsgBox, 308, Confirm Save, %saveFileName% already exists.`nDo you want to replace it?
			IfMsgBox, No
			{
				Return
			}
		}
		
		SaveFile := quoteMark
		SaveFile .= saveFilePath
		SaveFile .= quoteMark
		
		GuiControlGet, option, , RA1
		if (option)
		{	
			Run Convert.exe -append %FileNames% %SaveFile%,, UseErrorLevel Hide
		}
		else
		{
			Run Convert.exe +append %FileNames% %SaveFile%,, UseErrorLevel Hide
		}
		
		if (ErrorLevel == 0)
		{
			; play success sound
			SoundPlay *64
			
			; open saved location if it's currently not open
			SplitPath, filename, , SaveFileDir
			Dir := SaveFileDir
			StringSplit, array, Dir, \
			numArr := array0
			FolderName := array%numArr%
			
			IfWinNotExist, %FolderName%
			{
				Run, %SaveFileDir%, ,UseErrorLevel
			}
		}
		else
		{
			SoundPlay *16
			MSgBox, 0, Error, Failed to concatenate images!
		}
	}
	
	Return
}

Return

; -------------------------------------------------------------------------------------------
; Helper function to create groupbox inside Gui
; -------------------------------------------------------------------------------------------
GroupBox(GBvName			;Name for GroupBox control variable
		,Title				;Title for GroupBox
		,TitleHeight		;Height in pixels to allow for the Title
		,Margin				;Margin in pixels around the controls
		,Piped_CtrlvNames	;Pipe (|) delimited list of Controls
		,FixedWidth=""		;Optional fixed width
		,FixedHeight="")	;Optional fixed height
{
	Local maxX, maxY, minX, minY, xPos, yPos ;all else assumed Global
	minX:=99999
	minY:=99999
	maxX:=0
	maxY:=0
	Loop, Parse, Piped_CtrlvNames, |, %A_Space%
	{
		;Get position and size of each control in list.
		GuiControlGet, GB, Pos, %A_LoopField%
		;creates GBX, GBY, GBW, GBH
		if(GBX<minX) ;check for minimum X
			minX:=GBX
		if(GBY<minY) ;Check for minimum Y
			minY:=GBY
		if(GBX+GBW>maxX) ;Check for maximum X
			maxX:=GBX+GBW
		if(GBY+GBH>maxY) ;Check for maximum Y
			maxY:=GBY+GBH

		;Move the control to make room for the GroupBox
		xPos:=GBX+Margin
		yPos:=GBY+TitleHeight+Margin ;fixed margin
		GuiControl, Move, %A_LoopField%, x%xPos% y%yPos%
	}
	;re-purpose the GBW and GBH variables
	if(FixedWidth)
		GBW:=FixedWidth
	else
		GBW:=maxX-minX+2*Margin ;calculate width for GroupBox
	if(FixedHeight)
		GBH:=FixedHeight
	else
		GBH:=maxY-MinY+TitleHeight+2*Margin ;calculate height for GroupBox ;fixed 2*margin

	;Add the GroupBox
	Gui, Add, GroupBox, v%GBvName% x%minX% y%minY% w%GBW% h%GBH%, %Title%
	
	Return
}

; -------------------------------------------------------------------------------------------
; User closed Gui - exit program
; -------------------------------------------------------------------------------------------
GuiClose:
{
	ExitApp
}