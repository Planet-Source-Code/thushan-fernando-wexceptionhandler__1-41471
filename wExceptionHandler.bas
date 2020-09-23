Attribute VB_Name = "wExceptionHandler"

Option Explicit
'///////////////////////////////////////////////////////////////////////////////////////
'//  WebSoftware wExceptionHandler (Public Release)
'///////////////////////////////////////////////////////////////////////////////////////
'// NOTE: This code is being used in a commercial application, I have released it under
'//       the idea that it will help people:)
'//
'// In return we ask that you provide a link back to our site or put our name in the credits/about box of your app.
'//
'// Visit: http://www.wSoftware.biz | http://HotHTML3Beta.wSoftware.biz
'//
'// Any problems please email thushan@wsoftware.biz
'///////////////////////////////////////////////////////////////////////////////////////

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, ByVal LPEXCEPTION_RECORD As Long, ByVal lngBytes As Long)
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Enum enumExceptionType
    enumExceptionType_AccessViolation = &HC0000005
    enumExceptionType_DataTypeMisalignment = &H80000002
    enumExceptionType_Breakpoint = &H80000003
    enumExceptionType_SingleStep = &H80000004
    enumExceptionType_ArrayBoundsExceeded = &HC000008C
    enumExceptionType_FaultDenormalOperand = &HC000008D
    enumExceptionType_FaultDivideByZero = &HC000008E
    enumExceptionType_FaultInexactResult = &HC000008F
    enumExceptionType_FaultInvalidOperation = &HC0000090
    enumExceptionType_FaultOverflow = &HC0000091
    enumExceptionType_FaultStackCheck = &HC0000092
    enumExceptionType_FaultUnderflow = &HC0000093
    enumExceptionType_IntegerDivisionByZero = &HC0000094
    enumExceptionType_IntegerOverflow = &HC0000095
    enumExceptionType_PriviledgedInstruction = &HC0000096
    enumExceptionType_InPageError = &HC0000006
    enumExceptionType_IllegalInstruction = &HC000001D
    enumExceptionType_NoncontinuableException = &HC0000025
    enumExceptionType_StackOverflow = &HC00000FD
    enumExceptionType_InvalidDisposition = &HC0000026
    enumExceptionType_GuardPageViolation = &H80000001
    enumExceptionType_InvalidHandle = &HC0000008
    enumExceptionType_ControlCExit = &HC000013A
End Enum


Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

Private Type CONTEXT
    FltF0        As Double
    FltF1        As Double
    FltF2        As Double
    FltF3        As Double
    FltF4        As Double
    FltF5        As Double
    FltF6        As Double
    FltF7        As Double
    FltF8        As Double
    FltF9        As Double
    FltF10       As Double
    FltF11       As Double
    FltF12       As Double
    FltF13       As Double
    FltF14       As Double
    FltF15       As Double
    FltF16       As Double
    FltF17       As Double
    FltF18       As Double
    FltF19       As Double
    FltF20       As Double
    FltF21       As Double
    FltF22       As Double
    FltF23       As Double
    FltF24       As Double
    FltF25       As Double
    FltF26       As Double
    FltF27       As Double
    FltF28       As Double
    FltF29       As Double
    FltF30       As Double
    FltF31       As Double
    IntV0        As Double
    IntT0        As Double
    IntT1        As Double
    IntT2        As Double
    IntT3        As Double
    IntT4        As Double
    IntT5        As Double
    IntT6        As Double
    IntT7        As Double
    IntS0        As Double
    IntS1        As Double
    IntS2        As Double
    IntS3        As Double
    IntS4        As Double
    IntS5        As Double
    IntFp        As Double
    IntA0        As Double
    IntA1        As Double
    IntA2        As Double
    IntA3        As Double
    IntA4        As Double
    IntA5        As Double
    IntT8        As Double
    IntT9        As Double
    IntT10       As Double
    IntT11       As Double
    IntRa        As Double
    IntT12       As Double
    IntAt        As Double
    IntGp        As Double
    IntSp        As Double
    IntZero      As Double
    Fpcr         As Double
    SoftFpcr     As Double
    Fir          As Double
    Psr          As Long
    ContextFlags As Long
    Fill(4)      As Long
End Type
Private Type EXCEPTION_RECORD
    ExceptionCode                                        As Long
    ExceptionFlags                                       As Long
    pExceptionRecord                                     As Long
    ExceptionAddress                                     As Long
    NumberParameters                                     As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)   As Long
End Type

Private Type EXCEPTION_POINTERS
    pExceptionRecord     As EXCEPTION_RECORD
    ContextRecord        As CONTEXT
End Type

Public blnIsHandlerInstalled As Boolean
Public Sub HandleTheException(strException As String, strProcedure As String)
With frmException
    .txtException.Text = "Procedure: " & strProcedure & vbCrLf & "Date: " & Date & vbCrLf & "Time: " & Time & vbCrLf & strException
    .Show vbModal
    If Not .bContinue Then
        If .bAutoStart Then Debug.Print "Insert your AutoStart Procedure Here"        '// AutoStart Procedure Ommitted for the PSC Example as it doesnt apply to you
        End '// You shouldn't use End but I did for this example... always unload all your forms and use Set <frmName> = Nothing
    End If
End With
End Sub
'///////////////////////////////////////////////////////////////////////////////////////
'// METHOD 1
'///////////////////////////////////////////////////////////////////////////////////////
'// THIS WILL RAISE THE ERROR. If theres no error handling in your code where it occurs
'// then *toot do toot do toot*
'///////////////////////////////////////////////////////////////////////////////////////
'Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
'    On Error Resume Next
'
'    Dim ExceptionRecord As EXCEPTION_RECORD
'    Dim strExceptionDescriptiosn As String
'
'    ExceptionRecord = ExceptionPtrs.pExceptionRecord
'
'    Do Until ExceptionRecord.pExceptionRecord = 0
'        CopyMemory ExceptionRecord, ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
'    Loop
'
'    strExceptionDescriptiosn = GetExceptionDescription(ExceptionRecord.ExceptionCode)
'
'    On Error GoTo 0
'    Err.Raise vbObjectError, "ExceptionHandler", "Exception: " & strExceptionDescriptiosn & " [" & GetExceptionName(ExceptionRecord.ExceptionCode) & "]" & vbCrLf & "ExceptionAddress : " & ExceptionRecord.ExceptionAddress
'End Function


'///////////////////////////////////////////////////////////////////////////////////////
'// Method 2
'///////////////////////////////////////////////////////////////////////////////////////
'// This will handle it regardless of whether your Error handler is there or not.
'///////////////////////////////////////////////////////////////////////////////////////
Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long
    On Error Resume Next
    Dim Rec As EXCEPTION_RECORD, strException As String, AutoRestart As Boolean
    Rec = ExceptionPtrs.pExceptionRecord
    Do Until Rec.pExceptionRecord = 0
        CopyMemory Rec, Rec.pExceptionRecord, Len(Rec)
    Loop
    strException = GetExceptionDescription(Rec.ExceptionCode)
    frmException.txtException.Text = "Description: " & strException & vbCrLf & "Exception_Type: " & GetExceptionName(Rec.ExceptionCode) & vbCrLf & "Exception_Address: " & Rec.ExceptionAddress & vbCrLf & "Time/Date: " & Time & " [" & Date & "]"     '"Procedure: " & strProcedure & vbCrLf & "Line: " & strLineNumber & vbCrLf & "Date: " & Date & vbCrLf & "Time: " & Time & vbCrLf & "Exception: " & strExceptionDescriptiosn & " [" & GetExceptionName(ExceptionRecord.ExceptionCode) & "]" & vbCrLf & "ExceptionAddress : " & ExceptionRecord.ExceptionAddress
    frmException.Show vbModal
    AutoRestart = frmException.bAutoStart
    If Not frmException.bContinue Then
        Unload frmException
        'SaveWorkspace False
        If AutoRestart Then MsgBox "Insert your AutoStart Code here" '//Call ShellExecute(fMainForm.hWnd, "open", AppPath & "hothtml3.exe" & " ", AppPath & "crash.data", AppPath, 1)
        End
    End If
    ExceptionHandler = -1
End Function
Public Sub InstallExceptionHandler()
    On Error Resume Next
    If Not blnIsHandlerInstalled Then
        Call SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
        blnIsHandlerInstalled = True
    End If
End Sub



Public Function GetExceptionDescription(ExceptionType As enumExceptionType) As String
    On Error Resume Next
    
    Dim strDescription As String
  
    Select Case ExceptionType
        
        Case enumExceptionType_AccessViolation
            strDescription = "Access Violation"
        
        Case enumExceptionType_DataTypeMisalignment
            strDescription = "Data Type Misalignment"
        
        Case enumExceptionType_Breakpoint
            strDescription = "Breakpoint"
        
        Case enumExceptionType_SingleStep
            strDescription = "Single Step"
        
        Case enumExceptionType_ArrayBoundsExceeded
            strDescription = "Array Bounds Exceeded"
        
        Case enumExceptionType_FaultDenormalOperand
            strDescription = "Float Denormal Operand"
        
        Case enumExceptionType_FaultDivideByZero
            strDescription = "Divide By Zero"
        
        Case enumExceptionType_FaultInexactResult
            strDescription = "Floating Point Inexact Result"
        
        Case enumExceptionType_FaultInvalidOperation
            strDescription = "Invalid Operation"
        
        Case enumExceptionType_FaultOverflow
            strDescription = "Float Overflow"
        
        Case enumExceptionType_FaultStackCheck
            strDescription = "Float Stack Check"
        
        Case enumExceptionType_FaultUnderflow
            strDescription = "Float Underflow"
        
        Case enumExceptionType_IntegerDivisionByZero
            strDescription = "Integer Divide By Zero"
        
        Case enumExceptionType_IntegerOverflow
            strDescription = "Integer Overflow"
        
        Case enumExceptionType_PriviledgedInstruction
            strDescription = "Privileged Instruction"
        
        Case enumExceptionType_InPageError
            strDescription = "In Page Error"
        
        Case enumExceptionType_IllegalInstruction
            strDescription = "Illegal Instruction"
        
        Case enumExceptionType_NoncontinuableException
            strDescription = "Non Continuable Exception"
        
        Case enumExceptionType_StackOverflow
            strDescription = "Stack Overflow"
        
        Case enumExceptionType_InvalidDisposition
            strDescription = "Invalid Disposition"
        
        Case enumExceptionType_GuardPageViolation
            strDescription = "Guard Page Violation"
        
        Case enumExceptionType_InvalidHandle
            strDescription = "Invalid Handle"
        
        Case enumExceptionType_ControlCExit
            strDescription = "Control-C Exit"
        
        Case Else
            strDescription = "Unknown Exception Error"
    
    End Select
    
    GetExceptionDescription = strDescription
End Function


Public Function GetExceptionName(ExceptionType As enumExceptionType) As String
    On Error Resume Next
    
    Dim strDescription As String
  
    Select Case ExceptionType
        
        Case enumExceptionType_AccessViolation
            strDescription = "EXCEPTION_ACCESS_VIOLATION"
        
        Case enumExceptionType_DataTypeMisalignment
            strDescription = "EXCEPTION_DATATYPE_MISALIGNMENT"
        
        Case enumExceptionType_Breakpoint
            strDescription = "EXCEPTION_BREAKPOINT"
        
        Case enumExceptionType_SingleStep
            strDescription = "EXCEPTION_SINGLE_STEP"
        
        Case enumExceptionType_ArrayBoundsExceeded
            strDescription = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
        
        Case enumExceptionType_FaultDenormalOperand
            strDescription = "EXCEPTION_FLT_DENORMAL_OPERAND"
        
        Case enumExceptionType_FaultDivideByZero
            strDescription = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
        
        Case enumExceptionType_FaultInexactResult
            strDescription = "EXCEPTION_FLT_INEXACT_RESULT"
        
        Case enumExceptionType_FaultInvalidOperation
            strDescription = "EXCEPTION_FLT_INVALID_OPERATION"
        
        Case enumExceptionType_FaultOverflow
            strDescription = "EXCEPTION_FLT_OVERFLOW"
        
        Case enumExceptionType_FaultStackCheck
            strDescription = "EXCEPTION_FLT_STACK_CHECK"
        
        Case enumExceptionType_FaultUnderflow
            strDescription = "EXCEPTION_FLT_UNDERFLOW"
        
        Case enumExceptionType_IntegerDivisionByZero
            strDescription = "EXCEPTION_INT_DIVIDE_BY_ZERO"
        
        Case enumExceptionType_IntegerOverflow
            strDescription = "EXCEPTION_INT_OVERFLOW"
        
        Case enumExceptionType_PriviledgedInstruction
            strDescription = "EXCEPTION_PRIVILEGED_INSTRUCTION"
        
        Case enumExceptionType_InPageError
            strDescription = "EXCEPTION_IN_PAGE_ERROR"
        
        Case enumExceptionType_IllegalInstruction
            strDescription = "EXCEPTION_ILLEGAL_INSTRUCTION"
        
        Case enumExceptionType_NoncontinuableException
            strDescription = "EXCEPTION_NONCONTINUABLE_EXCEPTION"
        
        Case enumExceptionType_StackOverflow
            strDescription = "EXCEPTION_STACK_OVERFLOW"
        
        Case enumExceptionType_InvalidDisposition
            strDescription = "EXCEPTION_INVALID_DISPOSITION"
        
        Case enumExceptionType_GuardPageViolation
            strDescription = "EXCEPTION_GUARD_PAGE_VIOLATION"
        
        Case enumExceptionType_InvalidHandle
            strDescription = "EXCEPTION_INVALID_HANDLE"
        
        Case enumExceptionType_ControlCExit
            strDescription = "EXCEPTION_CONTROL_C_EXIT"
        
        Case Else
            strDescription = "Unknown"
    
    End Select
    
    GetExceptionName = strDescription
End Function


Public Sub RaiseAnException(ExceptionType As enumExceptionType)
    
    RaiseException ExceptionType, 0, 0, 0
End Sub

Public Sub UninstallExceptionHandler()
    On Error Resume Next
    
    If blnIsHandlerInstalled Then
        Call SetUnhandledExceptionFilter(0&)
        blnIsHandlerInstalled = False
    End If
End Sub
