Function f(a, b)
  If IsEmpty(a) Then
    a = 0
  End If
  If IsEmpty(b) Then
    b = 0
  End If
  c = a + b
  f = c
End Function

' Explicit type conversion to avoid implicit type coercion issues.
x = f(1, CInt(Empty))
y = f(CInt(Empty), 1)
z = f(CInt(Empty), CInt(Empty))

MsgBox "x = " & x & ", y = " & y & ", z = " & z
' Output: x = 1, y = 1, z = 0
