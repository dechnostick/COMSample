Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)>
<ComVisible(True)>
Public Class Class1
    ' Debug interface
    Public Function Test(ByVal str As String) As String
        Console.WriteLine("【{0}】", str)
        Return String.Format("Hello, {0}!!", str)
    End Function
End Class
