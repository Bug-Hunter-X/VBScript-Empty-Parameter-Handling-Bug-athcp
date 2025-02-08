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

x = f(1, Empty)
y = f(Empty, 1)
z = f(Empty, Empty)