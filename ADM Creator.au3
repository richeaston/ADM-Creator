#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Richard Easton
 Date: 			 22/04/2017

 Script Function:
	Group Policy ADM Creator.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

ADM_Creator()

Func ADM_Creator()
	;main gui
	Global $FrmADM = GUICreate("Group Policy ADM Creator Version 2", 673, 521, -1, -1)
	GUISetFont(8, 400, 0, "MS Gothic")

	$AD = GUICtrlCreateGroup(" ADM Details ", 8, 8, 417, 105)
	GUICtrlCreateLabel("ADM Name", 16, 28, 70, 20)
	$ADM = GUICtrlCreateInput("", 88, 24, 321, 19)

	GUICtrlCreateLabel("Class Type", 16, 58, 70, 20)
	$classtype = GUICtrlCreateCombo("Select Class Type", 88, 53, 321, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	GUICtrlSetData($classtype, "USER|MACHINE|BOTH", "Select Class Type")

	GUICtrlCreateLabel("Category", 16, 86, 70, 20)
	$Category = GUICtrlCreateInput("", 88, 80, 321, 19)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	$Policy = GUICtrlCreateGroup(" Policy Details ", 8, 120, 657, 321)
	$Policies = GUICtrlCreateEdit("", 17, 140, 633, 289, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
	GUICtrlSetData(-1, "")
	GUICtrlCreateGroup("", -99, -99, 1, 1)


	$Controls = GUICtrlCreateGroup(" Controls ", 8, 448, 657, 65)
		$Add = GUICtrlCreateButton("Add New Policy", 16, 464, 163, 33)
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x0066CC)

	$Preview = GUICtrlCreateButton("Preview Policy", 184, 464, 155, 33)
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x0066CC)
		GUICtrlSetState($Preview, $GUI_DISABLE)

	$Save = GUICtrlCreateButton("Save Policy", 344, 464, 155, 33)
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x008000)
		GUICtrlSetState($save, $GUI_DISABLE)

	$Reset = GUICtrlCreateButton("Reset All Inputs", 504, 464, 155, 33)
		GUICtrlSetColor(-1, 0xFFFFFF)
		GUICtrlSetBkColor(-1, 0x800000)
		GUICtrlSetState($reset, $GUI_DISABLE)

	GUICtrlCreateGroup("", -99, -99, 1, 1)

	GUISetState(@SW_SHOW, $FrmADM)

	;new policy gui initialised (but hidden)
	Local 	$New_Policy = GUICreate("New Policy", 480, 340, -1, -1, $WS_EX_MDICHILD )
			GUISetFont(8, 400, 0, "MS Gothic")
			GUICtrlCreateGroup(" New Policy ", 8, 8, 457, 257)
				GUICtrlCreateLabel("Policy", 16, 32, 40, 15)
				$PolicyName = GUICtrlCreateInput("", 103, 28, 353, 19)

				GUICtrlCreateLabel("KeyName", 16, 61, 46, 15)
				$KeyName = GUICtrlCreateInput("", 104, 56, 353, 19)

				GUICtrlCreateLabel("Explain", 16, 88, 46, 15)
				$Explain = GUICtrlCreateEdit("", 104, 85, 353, 81, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_WANTRETURN))

				GUICtrlCreateLabel("Value Name", 16, 180, 64, 15)
				$Value_Name = GUICtrlCreateInput("", 104, 175, 353, 19)

				GUICtrlCreateLabel("Value ON", 16, 211, 52, 15)
				$VON_Combo = GUICtrlCreateCombo("", 104, 205, 169, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
				Guictrlsetdata(-1, "NONE|DELETED|NUMERIC|STRING", "Please select")
				$VON_Input = GUICtrlCreateInput("", 280, 205, 177, 19)

				GUICtrlCreateLabel("Value OFF", 16, 239, 58, 15)
				$VOFF_Combo = GUICtrlCreateCombo("", 104, 232, 169, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
				Guictrlsetdata(-1, "NONE|DELETED|NUMERIC|STRING", "Please select")
				$Voff_Input = GUICtrlCreateInput("", 280, 232, 176, 19)
			GUICtrlCreateGroup("", -99, -99, 1, 1)

			$Commit = GUICtrlCreateButton("Commit Policy", 8, 272, 155, 33)
			GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1, 0x0066CC)

			$Reset_Child = GUICtrlCreateButton("Reset Form", 312, 272, 155, 33)
			GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1, 0x800000)



			While 1
				$nMsg = GUIGetMsg()
				Switch $nMsg
					Case $GUI_EVENT_CLOSE
						ExitLoop
					Case $Add
						;switch gui's
						GUISetState(@SW_SHOW, $New_Policy)
						GUISetState(@SW_HIDE, $FrmADM)
					Case $Commit
						;read policy entries
						$PN_Read = GUICtrlRead($PolicyName)
						$KN_Read = GuictrlRead($KeyName)
						$EX_Read = GuiCtrlRead($Explain)
						$VN_Read = GuictrlRead($Value_Name)
						$VOnC_Read = GuictrlRead($VON_Combo)
						$VONI_Read = GuictrlRead($VON_Input)
						$VOfC_Read = Guictrlread($VOFF_Combo)
						$VOfI_Read = GuictrlRead($VOFF_Input)
						;write policy to main gui
						Guictrlsetdata($Policies, _
						'	POLICY "' & $PN_Read & '"' & @CRLF & _
						'		KEYNAME "' & $KN_Read & '"' & @CRLF & _
						'		EXPLAIN "' &$EX_Read & '"' & @CRLF & _
						'		VALUENAME "' & $VN_Read & '"' & @CRLF & _
						'		VAULEON ' & $VOnC_Read & " " &  $VONI_Read & @CRLF & _
						'		VALUEOFF ' & $VOfC_Read & " " & $VOfI_Read & @CRLF & _
						'	END POLICY' & @CRLF _
						, $Policies)
						;clear new policy inputs
						Guictrlsetdata($PolicyName,"","")
						Guictrlsetdata($KeyName,"","")
						Guictrlsetdata($Explain,"","")
						Guictrlsetdata($Value_Name,"","")
						Guictrlsetdata($VON_Combo,"NONE|DELETED|NUMERIC|STRING", "Please select")
						Guictrlsetdata($VON_Input,"","")
						Guictrlsetdata($VOFF_Combo,"NONE|DELETED|NUMERIC|STRING", "Please select")
						Guictrlsetdata($VOFF_Input,"","")
						;switch gui's
						GUISetState(@SW_HIDE, $New_Policy)
						GUISetState(@SW_SHOW, $FrmADM)
						GUICtrlSetState($Preview, $GUI_ENABLE)
						GUICtrlSetState($Save, $GUI_ENABLE)
						GUICtrlSetState($reset, $GUI_ENABLE)

					case $Preview
						$ADMN_Read = Guictrlread($adm)
						$CT_read = Guictrlread($classtype)
						$cat_read = Guictrlread($Category)
						$pol_read = guictrlread($Policies)
						$pol_write = 'CLASS ' & $CT_read & @CRLF & @CRLF & 'CATEGORY "' & $cat_read & '"' & @crlf & $pol_read & 'END CATEGORY'
						$PrevFile = @ScriptDir & "\PREVIEWPOL.TXT"
						FileWrite($PrevFile, $pol_write)
						runwait("notepad.exe " & $PrevFile)
						FileDelete($PrevFile)
					Case $save
						$ADMN_Read = Guictrlread($adm)
						$CT_read = Guictrlread($classtype)
						$cat_read = Guictrlread($Category)
						$pol_read = guictrlread($Policies)
						if $ADMN_Read = "" Then
							Msgbox(48, "Warning!", "ADM Name is Blank, Please check and Correct")
						elseif $CT_read = "Please Select" or $CT_read = "" Then
							Msgbox(48, "Warning!", "ClassType is either Blank or incorrect, Please check and Correct")
						elseif $cat_read = "" Then
							Msgbox(48, "Warning!", "Category is Blank, Please check and Correct")
						elseif $pol_read = "" Then
							Msgbox(48, "Warning!", "Policies Section is Blank, Please check and Correct")
						Else
							$pol_write = 'CLASS ' & $CT_read & @CRLF & @CRLF & 'CATEGORY "' & $cat_read & '"' & @crlf & $pol_read & 'END CATEGORY'
							$ADMFile = @DesktopDir & "\" & $ADMN_Read & ".adm"
							FileWrite($ADMFile, $pol_write)
						EndIf

				EndSwitch
			WEnd
			; Delete the previous GUIs and all controls.
			GUIDelete($FrmADM)
			GUIDelete($New_Policy)


EndFunc


