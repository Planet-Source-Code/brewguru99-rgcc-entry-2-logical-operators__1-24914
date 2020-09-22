Attribute VB_Name = "modOperators"
Option Explicit

Enum Logical_Operator
  Operator_XOr
  Operator_OR
  Operator_And
  Operator_Imp
  Operator_Eqv
  Operator_Not
End Enum

Private NNsLoaded As Boolean

Private bnnXOr As clsBrewNN
Private bnnOr  As clsBrewNN
Private bnnAnd As clsBrewNN
Private bnnImp As clsBrewNN
Private bnnEqv As clsBrewNN
Private bnnNot As clsBrewNN


Public Function Logic_Oper(ByVal InValA As Byte, ByVal InValB As Byte, ByVal Operator As Logical_Operator) As Byte
  Select Case Operator
    Case Operator_XOr: Logic_Oper = rgccXor(InValA, InValB)
    Case Operator_OR: Logic_Oper = rgccOr(InValA, InValB)
    Case Operator_And: Logic_Oper = rgccAnd(InValA, InValB)
    Case Operator_Imp: Logic_Oper = rgccImp(InValA, InValB)
    Case Operator_Eqv: Logic_Oper = rgccEqv(InValA, InValB)
    Case Operator_Not: Logic_Oper = rgccNot(InValA)
  End Select
End Function


Private Sub LoadNNs()
  Dim x As Boolean
  
  On Error Resume Next
  x = False
  x = (FileLen(App.Path & "\bnnXOr.bnn") = FileLen(App.Path & "\bnnXOr.bnn"))
  On Error GoTo 0
  
  Set bnnXOr = New clsBrewNN
  Set bnnOr = New clsBrewNN
  Set bnnAnd = New clsBrewNN
  Set bnnImp = New clsBrewNN
  Set bnnEqv = New clsBrewNN
  Set bnnNot = New clsBrewNN
  
  If (x) Then
    bnnXOr.ImportNN App.Path & "\bnnXOr.bnn"
    bnnOr.ImportNN App.Path & "\bnnOr.bnn"
    bnnAnd.ImportNN App.Path & "\bnnAnd.bnn"
    bnnImp.ImportNN App.Path & "\bnnImp.bnn"
    bnnEqv.ImportNN App.Path & "\bnnEqv.bnn"
    bnnNot.ImportNN App.Path & "\bnnNot.bnn"
  Else
    Call CreateNNs
  End If
  NNsLoaded = True
End Sub

Private Sub CreateNNs()
  Dim i As Long
  
  'Init the NN's
  Randomize Timer
  bnnXOr.ConstructNN Array(2, 4, 1)
  bnnOr.ConstructNN Array(2, 4, 1)
  bnnAnd.ConstructNN Array(2, 4, 1)
  bnnImp.ConstructNN Array(2, 4, 1)
  bnnEqv.ConstructNN Array(2, 4, 1)
  bnnNot.ConstructNN Array(1, 1)
  
  'Train the NN's in all the possible inputs.
  For i = 1 To 250
    '0, 0
    bnnXOr.SetInput Array(0, 0)
    bnnOr.SetInput Array(0, 0)
    bnnAnd.SetInput Array(0, 0)
    bnnImp.SetInput Array(0, 0)
    bnnEqv.SetInput Array(0, 0)
    bnnNot.SetInput Array(0)
    
    bnnXOr.Train Array(0)
    bnnOr.Train Array(0)
    bnnAnd.Train Array(0)
    bnnImp.Train Array(1)
    bnnEqv.Train Array(1)
    bnnNot.Train Array(1)
    
    '0, 1
    bnnXOr.SetInput Array(0, 1)
    bnnOr.SetInput Array(0, 1)
    bnnAnd.SetInput Array(0, 1)
    bnnImp.SetInput Array(0, 1)
    bnnEqv.SetInput Array(0, 1)
    
    bnnXOr.Train Array(1)
    bnnOr.Train Array(1)
    bnnAnd.Train Array(0)
    bnnImp.Train Array(1)
    bnnEqv.Train Array(0)
    
    '1, 0
    bnnXOr.SetInput Array(1, 0)
    bnnOr.SetInput Array(1, 0)
    bnnAnd.SetInput Array(1, 0)
    bnnImp.SetInput Array(1, 0)
    bnnEqv.SetInput Array(1, 0)
    
    bnnXOr.Train Array(1)
    bnnOr.Train Array(1)
    bnnAnd.Train Array(0)
    bnnImp.Train Array(0)
    bnnEqv.Train Array(0)
    
    '1, 1
    bnnXOr.SetInput Array(1, 1)
    bnnOr.SetInput Array(1, 1)
    bnnAnd.SetInput Array(1, 1)
    bnnImp.SetInput Array(1, 1)
    bnnEqv.SetInput Array(1, 1)
    bnnNot.SetInput Array(1)
    
    bnnXOr.Train Array(0)
    bnnOr.Train Array(1)
    bnnAnd.Train Array(1)
    bnnImp.Train Array(1)
    bnnEqv.Train Array(1)
    bnnNot.Train Array(0)
  Next
  
  bnnXOr.ExportNN App.Path & "\bnnXOr.bnn"
  bnnOr.ExportNN App.Path & "\bnnOr.bnn"
  bnnAnd.ExportNN App.Path & "\bnnAnd.bnn"
  bnnImp.ExportNN App.Path & "\bnnImp.bnn"
  bnnEqv.ExportNN App.Path & "\bnnEqv.bnn"
  bnnNot.ExportNN App.Path & "\bnnNot.bnn"
End Sub

Private Function rgccXor(ByVal InValA As Byte, ByVal InValB As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnXOr.SetInput Array(Abs((InValA And (2 ^ i)) <> 0), Abs((InValB And (2 ^ i)) <> 0))
    rgccXor = rgccXor + ((2 ^ i) * Abs(bnnXOr.GetOutput(1) >= 0.5))
  Next
End Function

Private Function rgccOr(ByVal InValA As Byte, ByVal InValB As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnOr.SetInput Array(Abs((InValA And (2 ^ i)) <> 0), Abs((InValB And (2 ^ i)) <> 0))
    rgccOr = rgccOr + ((2 ^ i) * Abs(bnnOr.GetOutput(1) >= 0.5))
  Next
End Function

Private Function rgccAnd(ByVal InValA As Byte, ByVal InValB As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnAnd.SetInput Array(Abs((InValA And (2 ^ i)) <> 0), Abs((InValB And (2 ^ i)) <> 0))
    rgccAnd = rgccAnd + ((2 ^ i) * Abs(bnnAnd.GetOutput(1) >= 0.5))
  Next
End Function

Private Function rgccImp(ByVal InValA As Byte, ByVal InValB As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnImp.SetInput Array(Abs((InValA And (2 ^ i)) <> 0), Abs((InValB And (2 ^ i)) <> 0))
    rgccImp = rgccImp + ((2 ^ i) * Abs(bnnImp.GetOutput(1) >= 0.5))
  Next
End Function

Private Function rgccEqv(ByVal InValA As Byte, ByVal InValB As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnEqv.SetInput Array(Abs((InValA And (2 ^ i)) <> 0), Abs((InValB And (2 ^ i)) <> 0))
    rgccEqv = rgccEqv + ((2 ^ i) * Abs(bnnEqv.GetOutput(1) >= 0.5))
  Next
End Function

Private Function rgccNot(ByVal InValA As Byte) As Byte
  Dim i As Long
  
  If (Not NNsLoaded) Then Call LoadNNs
  For i = 0 To 7
    bnnNot.SetInput Array(Abs((InValA And (2 ^ i)) <> 0))
    rgccNot = rgccNot + ((2 ^ i) * Abs(bnnNot.GetOutput(1) >= 0.5))
  Next
End Function
