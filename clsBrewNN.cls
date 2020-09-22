VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBrewNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'BrewNN - A simple class encapsulating a Neural Network.
Const VERSION_INFO = "BrewNN v1.0.5       " 'Must be 20 characters long! (Exptected by ImportNN and ExportNN)

Private Type cDendrite
  ConnectionStrength  As Double
  ConnectionTarget    As Variant 'Expected array: 0=Layer, 1=Neuron
  WeightChange        As Double
End Type

Private Type cNeuron
  Dendrites()         As cDendrite
  DendriteCount       As Long
  Value               As Double
  BaseValue           As Double
  Error               As Double
End Type

Private Type cLayer
  Neurons()           As cNeuron
  NeuronCount         As Long
  AdjacentLayers      As Variant 'Expected array: 0=Layer
End Type

Private Type cMemoryBank
  Pointer             As cNeuron 'Value points to current slot
  Neurons()           As cNeuron
  NeuronCount         As Long
  'Slots()             As Double
  'SlotCount           As Long
End Type


'Internal Variables
Private Layers()      As cLayer
Private LayerCount    As Long
Private MemBanks()    As cMemoryBank
Private MemBankCount  As Long

'Property Value Containers
Private mLearningFactor As Double
Private mMomentumFactor As Double
Private mJitterFactor    As Double
Private mRefreshed      As Double


Private Sub Class_Initialize()
  ReDim Layers(0)
  LayerCount = 0
  ReDim MemBanks(0)
  MemBankCount = 0
  mLearningFactor = 1.5
  mMomentumFactor = 0.2
  mJitterFactor = 0
End Sub

Private Sub Class_Terminate()
  ReDim Layers(0)
  ReDim MemBanks(0)
End Sub



'LearningFactor: Amount of amplification to weight changes. For no amplification, set this value to 1 (1.5 default).
Public Property Get LearningFactor() As Double
  LearningFactor = mLearningFactor
End Property

Public Property Let LearningFactor(tmpValue As Double)
  mLearningFactor = tmpValue
End Property

'MomentumFactor: Factor of previous weight change is included in new weight change. For no momentum, set this to 0 (0.2 default).
Public Property Get MomentumFactor() As Double
  MomentumFactor = mMomentumFactor
End Property

Public Property Let MomentumFactor(tmpValue As Double)
  mMomentumFactor = tmpValue
End Property

'JitterFactor: Adds white (random) noise to the input values. For no noise, set this to 0 (default).
Public Property Get JitterFactor() As Double
  JitterFactor = mJitterFactor
End Property

Public Property Let JitterFactor(tmpValue As Double)
  mJitterFactor = tmpValue
End Property





Public Sub ConstructNN(Structure As Variant, Optional Memory As Variant = Null)
'Create a new NN with randomized values.
  Dim i        As Long
  Dim j        As Long
  Dim k        As Long
  Dim x        As Variant
  Dim mLayer   As cLayer
  Dim mMemBank As cMemoryBank
  
  'Set default values
  mLearningFactor = 1.5
  mMomentumFactor = 0.2
  mJitterFactor = 0
  
  'Create Layers
  For i = 0 To UBound(Structure)
    mLayer.AdjacentLayers = Array(i - 1)
    ReDim mLayer.Neurons(1 To Structure(i))
    mLayer.NeuronCount = Structure(i)
    For j = 1 To mLayer.NeuronCount
      mLayer.Neurons(j).BaseValue = GetRand
      mLayer.Neurons(j).Value = 0
      If i Then 'Not the first layer
        ReDim mLayer.Neurons(j).Dendrites(1 To Structure(i - 1))
        mLayer.Neurons(j).DendriteCount = Structure(i - 1)
        For k = 1 To Structure(i - 1)
          mLayer.Neurons(j).Dendrites(k).ConnectionStrength = GetRand
          mLayer.Neurons(j).Dendrites(k).ConnectionTarget = Array(1, i, k)
        Next
      End If
    Next
    AddLayer mLayer
  Next
  'Create Memory Banks
  If Not IsNull(Memory) Then
    For i = 0 To UBound(Memory)
      ReDim mMemBank.Neurons(1 To Memory(i)(1))
      mMemBank.NeuronCount = Memory(i)(1)
      'mMemBank.SlotCount = Memory(i)(0)
      'ReDim mMemBank.Slots(0, mMemBank.NeuronCount)
      'mMemBank.Pointer ********** Pointer init code goes here
      For j = 1 To mMemBank.NeuronCount
        mMemBank.Neurons(j).BaseValue = GetRand
        mMemBank.Neurons(j).Value = 0
        ReDim mMemBank.Neurons(j).Dendrites(1 To Structure(UBound(Structure) - 1))
        mMemBank.Neurons(j).DendriteCount = Structure(UBound(Structure) - 1)
        For k = 1 To Structure(UBound(Structure) - 1)
          mMemBank.Neurons(j).Dendrites(k).ConnectionStrength = GetRand
          mMemBank.Neurons(j).Dendrites(k).ConnectionTarget = Array(1, UBound(Structure), k)
        Next
      Next
      AddMemBank mMemBank
    Next
    'Connect second layer to Memory Banks
    If LayerCount > 1 Then
      For i = 1 To Layers(2).NeuronCount
        For j = 1 To MemBankCount
          For k = 1 To MemBanks(j).NeuronCount
            Layers(2).Neurons(i).DendriteCount = Layers(2).Neurons(i).DendriteCount + 1
            ReDim Preserve Layers(2).Neurons(i).Dendrites(1 To Layers(2).Neurons(i).DendriteCount)
            Layers(2).Neurons(i).Dendrites(Layers(2).Neurons(i).DendriteCount).ConnectionStrength = GetRand
            Layers(2).Neurons(i).Dendrites(Layers(2).Neurons(i).DendriteCount).ConnectionTarget = Array(2, j, k)
          Next
        Next
      Next
    End If
  End If
End Sub

Public Sub ImportNN(File As String)
'Retreives an NN from a file. Immedately destroys the existing NN (if any). So be sure you save the NN for backup before calling the import.
  Dim hDirectory As Long    'Pointer to the Directory
  Dim hProperties As Long   'Pointer to Properties Section
  Dim hLayers As Long       'Pointer to Layers Section
  Dim hLog As Long          'Pointer to Log Section
  Dim hCurLayer As Long     'Pointer to Current Layer (used in both Layers Section and MemBanks Section)
  Dim hCurNeuron As Long    'Pointer to Current Neuron (used in both Layers Section and MemBanks Section)
  Dim hCurDendrite As Long  'Pointer to Current Dendrite (used in both Layers Section and MemBanks Section)
  Dim mLayer   As cLayer
  Dim a As String
  Dim i As Long, j As Long, k As Long, lngTmp As Long, TmpLayerCount As Long
  Dim FileNumb As Integer
  
  On Error Resume Next
  Call Class_Initialize
  If (FileLen(File) = FileLen(File)) Then
    FileNumb = FreeFile
    Open File For Binary As FileNumb
      a = String(20, " ")
      Get FileNumb, , a                               'Retreive version info
      If (a <> VERSION_INFO) Then                     'Exit if not the same version (no backward compability, sorry.)
        Close FileNumb
        Exit Sub
      End If
      
      'Directory
      '#########
      hDirectory = 282                                'Skip ahead to the directory (ignore the Human-Readable layer description)
      Seek FileNumb, hDirectory
      Get FileNumb, , j                               'Get the number of entries in the directory
      For i = 1 To j                                  'Check directory entries until Layers section found
        a = String(255, " ")
        Get FileNumb, , lngTmp
        Get FileNumb, , a
        Select Case a
          Case "Layers" & String(249, " ")
              hLayers = lngTmp
          Case "Properties" & String(245, " ")
              hProperties = lngTmp
          Case "Log" & String(252, " ")
              hLog = lngTmp
        End Select
      Next
      
      'Properties
      '########
      Seek FileNumb, hProperties
      Get FileNumb, , mLearningFactor
      Get FileNumb, , mMomentumFactor
      Get FileNumb, , mJitterFactor
      Get FileNumb, , mRefreshed
      
      ' Layers
      '########
      Seek FileNumb, hLayers
      Get FileNumb, , TmpLayerCount                      'Number of layers
      For i = 1 To TmpLayerCount
        Seek FileNumb, hLayers + 4 + (4 * (i - 1))        'Get layer position
        Get FileNumb, , hCurLayer
        Seek FileNumb, hCurLayer                      'Move to layer
        Get FileNumb, , mLayer.NeuronCount            'Create neurons for layer
        ReDim mLayer.Neurons(1 To mLayer.NeuronCount)
        For j = 1 To mLayer.NeuronCount               'Process neurons
          Seek FileNumb, hCurLayer + 4 + (4 * (j - 1))
          Get FileNumb, , hCurNeuron                  'Saved neuron position
          Seek FileNumb, hCurNeuron
          Get FileNumb, , mLayer.Neurons(j).DendriteCount     'Dendrites
          ReDim mLayer.Neurons(j).Dendrites(1 To mLayer.Neurons(j).DendriteCount)
          For k = 1 To mLayer.Neurons(j).DendriteCount
            Seek FileNumb, hCurNeuron + 4 + (4 * (k - 1))
            Get FileNumb, , hCurDendrite                      'Saved Dendrite position
            Seek FileNumb, hCurDendrite
            Get FileNumb, , mLayer.Neurons(j).Dendrites(k).WeightChange
            Get FileNumb, , mLayer.Neurons(j).Dendrites(k).ConnectionStrength  'Saved Dendrite information
            mLayer.Neurons(j).Dendrites(k).ConnectionTarget = Array(0, 0, 0)
            Get FileNumb, , lngTmp
            mLayer.Neurons(j).Dendrites(k).ConnectionTarget(0) = lngTmp
            Get FileNumb, , lngTmp
            mLayer.Neurons(j).Dendrites(k).ConnectionTarget(1) = lngTmp
            Get FileNumb, , lngTmp
            mLayer.Neurons(j).Dendrites(k).ConnectionTarget(2) = lngTmp
          Next
          Seek FileNumb, hCurNeuron + 4 + (4 * mLayer.Neurons(j).DendriteCount) 'Saved Neuron information
          Get FileNumb, , mLayer.Neurons(j).Value
          Get FileNumb, , mLayer.Neurons(j).BaseValue
          Get FileNumb, , mLayer.Neurons(j).Error
        Next
        Seek FileNumb, hCurLayer + 4 + (4 * mLayer.NeuronCount)
        Get FileNumb, , lngTmp
        ReDim mLayer.AdjacentLayers(0 To lngTmp - 1)  'Saved layer information
        For j = 0 To UBound(mLayer.AdjacentLayers)
          Get FileNumb, , lngTmp
          mLayer.AdjacentLayers(j) = lngTmp
        Next
        AddLayer mLayer
      Next
    Close FileNumb
  End If
End Sub

Public Sub ExportNN(File As String)
'Saves the NN to a file. I avoided taking a few shortcuts here, keeping future expansion in mind.
  Dim hDirectory As Long    'Pointer to the Directory
  Dim hProperties As Long   'Pointer to Properties Section
  Dim hLayers As Long       'Pointer to Layers Section
  Dim hLog As Long          'Pointer to Log Section
  Dim hCurLayer As Long     'Pointer to Current Layer (used in both Layers Section and MemBanks Section)
  Dim hCurNeuron As Long    'Pointer to Current Neuron (used in both Layers Section and MemBanks Section)
  Dim hCurDendrite As Long  'Pointer to Current Dendrite (used in both Layers Section and MemBanks Section)
  Dim a As String
  Dim i As Long, j As Long, k As Long
  Dim FileNumb As Integer
  
  On Error Resume Next
  If (FileLen(File) = FileLen(File)) Then Kill (File)
  FileNumb = FreeFile
  Open File For Binary As FileNumb
    a = VERSION_INFO
    a = a & String(20 - Len(a), " ") & vbCrLf
    Put FileNumb, , a '20 Char Version Info followed by carrage return
    a = "Layer Structure: "
    For i = 1 To LayerCount                         '255 Char Human-Readable description of layer structure followed by double carrage return
      If (i > 1) Then a = a & ", "
      a = a & CStr(Layers(i).NeuronCount)
    Next
    a = a & String(255 - Len(a), " ") & vbCrLf & vbCrLf
    Put FileNumb, , a
    
    'Directory
    '#########
    'Each Directory entry is a pointer followed by a 255 Char description of the section
    hDirectory = Seek(FileNumb)                     'Save Directory Position (Should be = 282)
    Put FileNumb, , 3&                              'Number of entries in directory
    Put FileNumb, , 0&                              'Pointer to Properties
    a = "Properties": a = a & String(255 - Len(a), " ")
    Put FileNumb, , a
    Put FileNumb, , 0&                              'Pointer to Layers (Set to 0 until position determined later)
    a = "Layers": a = a & String(255 - Len(a), " ")
    Put FileNumb, , a
    Put FileNumb, , 0&                               'Pointer to Log
    a = "Log": a = a & String(255 - Len(a), " ")
    Put FileNumb, , a
    
    
    'Properties
    '########
    hProperties = Seek(FileNumb)
    Seek FileNumb, hDirectory + 4
    Put FileNumb, , hProperties
    Seek FileNumb, hProperties
    Put FileNumb, , mLearningFactor
    Put FileNumb, , mMomentumFactor
    Put FileNumb, , mJitterFactor
    Put FileNumb, , mRefreshed

    ' Layers
    '########
    hLayers = Seek(FileNumb)                        'Save Layer Position
    Seek FileNumb, hDirectory + 263
    Put FileNumb, , hLayers
    Seek FileNumb, hLayers
    Put FileNumb, , LayerCount                      'Number of layers
    For i = 1 To LayerCount
      Put FileNumb, , 0&                            'Pointers to each layer (Set to 0 until position determined later)
    Next
    For i = 1 To LayerCount
      hCurLayer = Seek(FileNumb)                    'Save Current Layer Position
      Seek FileNumb, hLayers + 4 + (4 * (i - 1))
      Put FileNumb, , hCurLayer
      Seek FileNumb, hCurLayer
      Put FileNumb, , Layers(i).NeuronCount         'Number of neurons in current layer
      For j = 1 To Layers(i).NeuronCount
        Put FileNumb, , 0&                          'Pointers to neurons in current layer (Set to 0 until position determined later)
      Next
      Put FileNumb, , CLng(UBound(Layers(i).AdjacentLayers) + 1) 'Number of adjacent layers
      For j = 0 To UBound(Layers(i).AdjacentLayers)              'Save Adjacent Layer info
        Put FileNumb, , CLng(Layers(i).AdjacentLayers(j))
      Next
      For j = 1 To Layers(i).NeuronCount            'Process neurons in current layer
        hCurNeuron = Seek(FileNumb)                         'Save neuron position
        Seek FileNumb, hCurLayer + 4 + (4 * (j - 1))
        Put FileNumb, , hCurNeuron
        Seek FileNumb, hCurNeuron
        Put FileNumb, , Layers(i).Neurons(j).DendriteCount  'Number of dendrites in current neuron
        For k = 1 To Layers(i).Neurons(j).DendriteCount
          Put FileNumb, , 0&                                'Pointers to dendrites in current neuron (Set to 0 until position determined later)
        Next
        Put FileNumb, , Layers(i).Neurons(j).Value
        Put FileNumb, , Layers(i).Neurons(j).BaseValue
        Put FileNumb, , Layers(i).Neurons(j).Error
        For k = 1 To Layers(i).Neurons(j).DendriteCount     'Process dendrites in current neuron
          hCurDendrite = Seek(FileNumb)
          Seek FileNumb, hCurNeuron + 4 + (4 * (k - 1))
          Put FileNumb, , hCurDendrite                      'Save dendrite position
          Seek FileNumb, hCurDendrite
          Put FileNumb, , Layers(i).Neurons(j).Dendrites(k).WeightChange
          Put FileNumb, , Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Put FileNumb, , CLng(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0))
          Put FileNumb, , CLng(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1))
          Put FileNumb, , CLng(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2))
        Next
      Next
    Next
  Close FileNumb
End Sub

Private Sub AddLayer(Layer As cLayer)
'Used by ConstructNN
  LayerCount = LayerCount + 1
  ReDim Preserve Layers(0 To LayerCount)
  Layers(LayerCount) = Layer
End Sub

Private Sub AddMemBank(MemBank As cMemoryBank)
'Used by ConstructNN
  MemBankCount = MemBankCount + 1
  ReDim Preserve MemBanks(0 To MemBankCount)
  MemBanks(MemBankCount) = MemBank
End Sub




Public Sub SetInput(Data As Variant)
  Dim i As Long
  
  If LayerCount Then
    For i = 1 To Layers(1).NeuronCount
      Layers(1).Neurons(i).Value = CDbl(Data(i - 1)) + (1 * mJitterFactor * GetRand)
    Next
  End If
  mRefreshed = False
End Sub

Public Sub Refresh()
'Update output values for current input.
  Dim a As Variant, b As Variant
  Dim i As Long, j As Long, k As Long
  Dim x As Variant
  
  'Avoid redundant refreshes
  If (mRefreshed) Then
    Exit Sub
  Else
    mRefreshed = True
    
    'Layers
    For i = 2 To LayerCount
      For j = 1 To Layers(i).NeuronCount
        DoEvents
        'Value =  SigmoidSquash(With:Dendrites(Value + ConnectionStrength * TargetsValue) + BaseValue)
        For k = 1 To Layers(i).Neurons(j).DendriteCount
          Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
            Case 1 'Target is a layer
                x = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
            Case 2 'Target is a MemBank
                x = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
            Case Else
                x = 0
          End Select
          Layers(i).Neurons(j).Value = Layers(i).Neurons(j).Value + x * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
        Next
        Layers(i).Neurons(j).Value = SigmoidSquash(Layers(i).Neurons(j).Value + Layers(i).Neurons(j).BaseValue)
      Next
    Next
    'MemBanks
    For i = 1 To MemBankCount
      For j = 1 To MemBanks(i).NeuronCount
        DoEvents
        'Value =  SigmoidSquash(FOR:Dendrites(Value + ConnectionStrength * TargetsValue) + BaseValue)
        For k = 1 To MemBanks(i).Neurons(j).DendriteCount
          Select Case MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
            Case 1 'Target is a Layer
                x = Layers(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
            Case 2 'Target is a MemBank
                x = MemBanks(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
            Case Else
                x = 0
          End Select
          MemBanks(i).Neurons(j).Value = MemBanks(i).Neurons(j).Value + x * MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength
        Next
        MemBanks(i).Neurons(j).Value = SigmoidSquash(MemBanks(i).Neurons(j).Value + MemBanks(i).Neurons(j).BaseValue)
      Next
      '**********Code for slot save and pointer adjustment goes here
    Next
  End If
End Sub

Public Function GetOutput(Optional Index As Long = 0) As Variant
  'Return specific output neuron or entire output level (default)
  Dim i As Long
  Dim x As Variant
  
  'Refresh if necessary
  If (Not mRefreshed) Then Me.Refresh
  
  If Index Then
    GetOutput = CDbl(Layers(LayerCount).Neurons(Index).Value)
  Else
    ReDim Preserve x(1 To Layers(LayerCount).NeuronCount)
    For i = 1 To Layers(LayerCount).NeuronCount
      x(i) = CDbl(Layers(LayerCount).Neurons(i).Value)
    Next
    GetOutput = x
  End If
End Function

Public Sub Train(Target As Variant)
'Determine error levels for expected (target) output and actual output; then adjust the connection weights
  Dim i As Long, j As Long, k As Long
  Dim TargetValue  As Double
  Dim WeightChange As Double
  
  'Refresh if necessary
  If (Not mRefreshed) Then Me.Refresh
  
  'Determine Error Level
  For i = 1 To Layers(LayerCount).NeuronCount
    Layers(LayerCount).Neurons(i).Error = CDbl(Target(i - 1)) - Layers(LayerCount).Neurons(i).Value 'Raw Error
  Next
  
  'Back Propigate Errors
  For i = LayerCount To 2 Step -1
    For j = 1 To Layers(i).NeuronCount
      DoEvents
      'Determine Neuron Error
      Layers(i).Neurons(j).Error = Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Value * (1 - Layers(i).Neurons(j).Value)
      For k = 1 To Layers(i).Neurons(j).DendriteCount 'Determine Dendrite Error
        Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1
              Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error + Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Case 2
              MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Error + Layers(i).Neurons(j).Error * Layers(i).Neurons(j).Dendrites(k).ConnectionStrength
          Case Else
              'Ignore
        End Select
      Next
    Next
  Next
  
  'Update Layer Weights
  For i = 2 To LayerCount
    For j = 1 To Layers(i).NeuronCount
      DoEvents
      For k = 1 To Layers(i).Neurons(j).DendriteCount
        Select Case Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1 'Target is a Layer
              TargetValue = Layers(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2 'Target is a MemBank
              TargetValue = MemBanks(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(Layers(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              TargetValue = 0
        End Select
        'ConectionStrength = ConnectionStrength + mLearningFactor * TargetsValue * NeuronError + mMomentumFactor * WeightChange
        WeightChange = mLearningFactor * TargetValue * Layers(i).Neurons(j).Error + mMomentumFactor * Layers(i).Neurons(j).Dendrites(k).WeightChange
        Layers(i).Neurons(j).Dendrites(k).WeightChange = WeightChange
        Layers(i).Neurons(j).Dendrites(k).ConnectionStrength = Layers(i).Neurons(j).Dendrites(k).ConnectionStrength + WeightChange
      Next
      Layers(i).Neurons(j).BaseValue = Layers(i).Neurons(j).BaseValue + mLearningFactor * Layers(i).Neurons(j).Error
      Layers(i).Neurons(j).Error = 0
    Next
  Next
  'Update MemBank Weights
  For i = 2 To MemBankCount
    For j = 1 To MemBanks(i).NeuronCount
      DoEvents
      For k = 1 To MemBanks(i).Neurons(j).DendriteCount
        Select Case MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(0)
          Case 1 'Target is a Layer
              TargetValue = Layers(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case 2 'Target is a MemBank
              TargetValue = MemBanks(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(1)).Neurons(MemBanks(i).Neurons(j).Dendrites(k).ConnectionTarget(2)).Value
          Case Else
              TargetValue = 0
        End Select
        'ConectionStrength = ConnectionStrength + mLearningFactor * TargetsValue * NeuronError + mMomentumFactor * WeightChange
        WeightChange = mLearningFactor * TargetValue * MemBanks(i).Neurons(j).Error + mMomentumFactor * MemBanks(i).Neurons(j).Dendrites(k).WeightChange
        MemBanks(i).Neurons(j).Dendrites(k).WeightChange = WeightChange
        MemBanks(i).Neurons(j).Dendrites(k).ConnectionStrength = MemBanks(i).Neurons(j).Value + WeightChange
      Next
      MemBanks(i).Neurons(j).BaseValue = MemBanks(i).Neurons(j).BaseValue + mLearningFactor * MemBanks(i).Neurons(j).Error
      MemBanks(i).Neurons(j).Error = 0
    Next
  Next
End Sub





Private Function SigmoidSquash(InVal As Double) As Double
  'Uses Sigmoid Function to squash InVal into interval 0..1
  On Error GoTo SigmoidSquash_Error
  Select Case InVal
    Case Is > 50:  SigmoidSquash = 1
    Case Is < -50: SigmoidSquash = 0
    Case Else:     SigmoidSquash = 1 / (1 + Exp(-InVal))
  End Select
Exit Function
SigmoidSquash_Error:
  SigmoidSquash = 0
  Err.Clear
End Function

Private Function GetRand() As Double
'Produces random number between -1 and 1
  Do
    GetRand = 1 - Rnd - Rnd
  Loop Until Abs(GetRand) >= 0.01
End Function

Public Sub Jitter(Optional MaxVariance As Double = 0.05)
'Helps avoid hill-descent problems by slightly randomizing TransferWeight
  Dim i As Long, j As Long, k As Long
  
  For i = 2 To LayerCount
    For j = 1 To Layers(i).NeuronCount
      For k = 1 To Layers(i).Neurons(j).DendriteCount
        Layers(i).Neurons(j).Dendrites(k).ConnectionStrength = Layers(i).Neurons(j).Dendrites(k).ConnectionStrength + (MaxVariance * GetRand)
      Next
    Next
  Next
End Sub
'W_i = (w1_i) + (w2_i) + ... (w#_i)
'Weight Decay
'(w_i+1) = (w_i)^2 + (
'Weight Elimination
'(w_i+1) = (w_i)^2/(w_i) ^ 2 + EliminationConstant ^ 2

