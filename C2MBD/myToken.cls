VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum eTokenType
    tkInvalid
    tkWord
    
    tkAssign
    tkRBR
    tkLBR
    tkColon
    tkSemi
    tkComma
    tkLCBR
    tkRCBR
    tkAdd
    tkSub
    tkMul
    tkDiv
    tkMod
    tkGT
    tkLT
    tkLSBR
    tkRSBR
    tkBitOr
    tkBitAnd
    tkBitXor
    tkBitNot
    tkNOT

    tkRShift
    tkLShift
    tkInc
    tkDec
    tkGTE
    tkLTE
    tkOr
    tkAnd
    tkEQU
    tkNEQ
    tkType
    tkStatic
    tkReturn
    tkElse
    tkIf
    tkSw
    tkCase
    tkDefault
    tkBreak

End Enum

Public mTxt As String
Public mType As eTokenType
