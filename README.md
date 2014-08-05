COMSample
=========

VB を COM で C から呼び出す

```VB.net
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)>
<ComVisible(True)>
Public Class Class1
    ' Debug interface
    Public Function Test(ByVal str As String) As String
        Console.WriteLine("[{0}]", str)
        Return String.Format("Hello, {0}!!", str)
    End Function
End Class
```

を

```c
pIClass1->lpVtbl->Test((void*)pIClass1, bstr_in, &bstr_out);
```

な感じで呼び出す。
呼び出されると以下のような出力を得る。

```dos
[Test]
Hello, Test!!
```

vbs で呼び出すとこのようになる。


