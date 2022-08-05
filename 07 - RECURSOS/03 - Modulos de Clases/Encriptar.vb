Imports System
Imports System.Security.Cryptography
Public Class ENCRIPTAR
    Public Shared Function ENCRIPTAR_PALABRA(ByVal VALOR_ORIGINAL As String) As String
        Dim _md5 As MD5 = MD5.Create()
        Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(VALOR_ORIGINAL)
        Dim hash As Byte() = _md5.ComputeHash(inputBytes)
        Return BitConverter.ToString(hash).Replace("-", "")
    End Function
End Class
