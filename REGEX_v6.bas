Attribute VB_Name = "REGEX_v6"
Option Explicit

    
   'This sub to display in the debug window what functions here return
    Public Sub Example()
        Const Separator As String = "|"
        Dim TestString As String: TestString = "First Name|Last Name|Phone Number|Country|Site"
        
        Debug.Print "Test string: " & TestString
        Debug.Print "Separator: " & Separator & vbCrLf
        
        Debug.Print "Position of the second separator within the test string: " & GetStrPos(TestString, Separator, 1, 2)
        Debug.Print "Retrieve the value of the field number 0 within the test string: " & GetField(TestString, 0, Separator)
        Debug.Print "Retrieve the value of the field number 1 within the test string: " & GetField(TestString, 1, Separator)
        Debug.Print "Retrieve the value of the field number 2 within the test string: " & GetField(TestString, 2, Separator)
        Debug.Print "Retrieve the value of the field number 3 within the test string: " & GetField(TestString, 3, Separator)
        Debug.Print "Retrieve the value of the field number 4 within the test string: " & GetField(TestString, 4, Separator)
        
        Debug.Print "Modify the whole record:"
        LetField TestString, 0, Separator, "Pat", vbTextCompare
        LetField TestString, 1, Separator, "V", vbTextCompare
        LetField TestString, 2, Separator, "+33", vbTextCompare
        LetField TestString, 3, Separator, "France", vbTextCompare
        LetField TestString, 4, Separator, "Paris", vbTextCompare
        Debug.Print TestString
        
    End Sub

   
   
   'Retrieve the value of a field from a record
   'Fields start from 0
    Public Property Get GetField( _
                                 ByRef Record As String, _
                                 ByRef FieldID As Long, _
                                 ByRef Separator As String _
                               ) As String
        
        If Record = "" Then Exit Property
        Dim Nbr As Long: Nbr = GetStrOcc(Record, Separator)
        Dim n1 As Long, n2 As Long, n3 As Long

        Select Case FieldID
            Case Is > Nbr
                GetField = ""
            Case Nbr
                GetField = Right(Record, Len(Record) - GetStrPos(Record, Separator, , FieldID) - (Len(Separator) - 1))
            Case 0
                If Nbr = 0 Then
                    GetField = Record
                    Else
                        GetField = Left(Record, GetStrPos(Record, Separator, 1) - 1)
                End If
            Case Else
                n1 = GetStrPos(Record, Separator, 1, FieldID)
                n2 = Len(Separator)
                n3 = GetStrPos(Record, Separator, n1 + n2, 1)
                GetField = Mid(Record, n1 + n2, n3 - n1 - n2)
        End Select
    End Property
    
   'Create/modify a value in a record
   'Fields start from 0
    Public Sub LetField( _
                                 ByRef Record As String, _
                                 ByRef FieldID As Long, _
                                 ByRef Separator As String, _
                                 ByRef NewValue As String, _
                                 Optional ByRef CompareMethod As VbCompareMethod = vbTextCompare _
                                 )
        
        
        If Record = "" Then Exit Sub
        Dim Nbr As Long: Nbr = GetStrOcc(Record, Separator, 1, CompareMethod)
        Dim n1 As Long, n2 As Long, n3 As Long
        Select Case FieldID
            Case Is > Nbr
                Exit Sub
            Case Nbr
                Record = Replace(Record, Right(Record, Len(Record) - GetStrPos(Record, Separator, , FieldID, CompareMethod) - (Len(Separator) - 1)), NewValue, 1, 1, CompareMethod)
            Case 0
                If Nbr = 0 Then
                    Record = NewValue
                    Else
                        Record = Replace(Record, Left(Record, GetStrPos(Record, Separator, 1, 1, CompareMethod) - 1), NewValue, 1, 1, CompareMethod)
                End If
            Case Else
                n1 = GetStrPos(Record, Separator, 1, FieldID, CompareMethod)
                n2 = Len(Separator)
                n3 = GetStrPos(Record, Separator, n1 + n2, 1, CompareMethod)
                Record = Replace(Record, Mid(Record, n1 + n2, n3 - n1 - n2), NewValue, 1, 1, CompareMethod)
        End Select
    End Sub




   'Number of occurrences of a string within another string
    Private Property Get GetStrOcc( _
                                    ByRef SearchIn As String, _
                                    ByRef SearchFor As String, _
                                    Optional ByVal StartFrom As Long = 1, _
                                    Optional ByRef CompareMethod As VbCompareMethod = vbTextCompare _
                                  ) As Long
        
       'Use some Statics in order to not reevaluate the property during multiple calls with same parameters
        Static sSearchIn As String
        Static sSearchFor As String
        Static sStartFrom As Long
        Static sCompareMethod As VbCompareMethod
        Static sGetStrOcc As Long
        If SearchIn = sSearchIn And _
           SearchFor = sSearchFor And _
           StartFrom = sStartFrom And _
           CompareMethod = sCompareMethod Then
                GetStrOcc = sGetStrOcc
                Exit Property
                Else
                    sSearchIn = SearchIn
                    sSearchFor = SearchFor
                    sStartFrom = StartFrom
                    sCompareMethod = CompareMethod
                    sGetStrOcc = 0
        End If
        
       'Evaluate the property
        Dim lng As Long: lng = Len(SearchFor)
        Dim tmp As Long
        tmp = InStr(StartFrom, SearchIn, SearchFor, CompareMethod)
        If tmp > 0 Then
            Do
                sGetStrOcc = sGetStrOcc + 1
                tmp = InStr(tmp + lng, SearchIn, SearchFor, CompareMethod)
            Loop While tmp <> 0
        End If
        GetStrOcc = sGetStrOcc
    End Property


   'Position of a string X within a string Y,
    Private Property Get GetStrPos( _
                                  ByRef SearchIn As String, _
                                  ByRef SearchFor As String, _
                                  Optional ByVal StartFrom As Long = 1, _
                                  Optional ByRef Occurence As Long = 1, _
                                  Optional ByRef CompareMethod As VbCompareMethod = vbTextCompare _
                                  ) As Long
        
        
       'Use some Statics in order to not reevaluate the property during multiple calls with same parameters
        Static sSearchIn As String
        Static sSearchFor As String
        Static sStartFrom As Long
        Static sOccurence As Long
        Static sCompareMethod As VbCompareMethod
        Static sGetStrPos As Long
        If SearchIn = sSearchIn And _
           SearchFor = sSearchFor And _
           StartFrom = sStartFrom And _
           Occurence = sOccurence And _
           CompareMethod = sCompareMethod Then
                GetStrPos = sGetStrPos
                Exit Property
                Else
                    sSearchIn = SearchIn
                    sSearchFor = SearchFor
                    sStartFrom = StartFrom
                    sOccurence = Occurence
                    sCompareMethod = CompareMethod
                    sGetStrPos = 0
        End If
        
       'Evaluate the property
        Dim lng As Long: lng = Len(SearchFor)
        Dim tmp As Long: tmp = InStr(StartFrom, SearchIn, SearchFor, CompareMethod)
        Dim Nbr As Long
        If tmp > 0 Then
            Nbr = Nbr + 1
            Do While Nbr < Occurence
                tmp = InStr(tmp + lng, SearchIn, SearchFor, CompareMethod)
                Nbr = Nbr + 1
            Loop
        End If
        sGetStrPos = tmp
        GetStrPos = tmp
    End Property


