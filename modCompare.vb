Module modCompare
    Public Function lt(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("Less Than")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) < CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) < CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) < CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) < CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function

    Public Function lteq(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("Less Than Or Equal")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) <= CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) <= CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) <= CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) <= CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function

    Public Function mteq(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("More Than Or Equal")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) >= CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) >= CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) >= CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) >= CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function

    Public Function neq(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("Not Equal")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) <> CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) <> CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) <> CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) <> CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function

    Public Function mt(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("More Than")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) > CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) > CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) > CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) > CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function

    Public Function eq(ByVal foValue1 As Object, ByVal foValue2 As Object) As Boolean
        Dim typeCode As TypeCode = Type.GetTypeCode(foValue1.GetType())

        On Error GoTo errProc
        Debug.Print("Equal")

        Select Case typeCode
            Case typeCode.Boolean
                Console.WriteLine("Boolean: {0}", foValue1)
                Console.WriteLine("Boolean: {0}", foValue2)
                Return CBool(foValue1) = CBool(foValue2)
            Case typeCode.Double
                Console.WriteLine("Double: {0}", foValue1)
                Console.WriteLine("Double: {0}", foValue2)
                Return CDbl(foValue1) = CDbl(foValue2)
            Case typeCode.Int16, typeCode.Int32, typeCode.Int64
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CInt(foValue1) = CInt(foValue2)
            Case Else
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue1)
                Console.WriteLine("{0}: {1}", typeCode.ToString(), foValue2)
                Return CStr(foValue1) = CStr(foValue2)
        End Select

errProc:
        'on conversion error RETURN FALSE
        Return False
    End Function
End Module
