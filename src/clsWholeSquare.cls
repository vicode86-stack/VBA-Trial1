VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWholeSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pA As Double
Private pB As Double

Public Property Let A(value As Double)
    pA = value
End Property

Public Property Get A() As Double
    A = pA
End Property

Public Property Let B(value As Double)
    pB = value
End Property

Public Property Get B() As Double
    B = pB
End Property

Public Function Calculate() As Double
    Calculate = pA ^ 2 + 2 * pA * pB + B ^ 2
End Function

