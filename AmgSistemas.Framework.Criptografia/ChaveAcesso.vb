Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace SNGenerator

    Public Enum SNKeyLength

        SN16 = 16
        SN20 = 20
        SN24 = 24
        SN28 = 28
        SN32 = 32

    End Enum

    Public Enum SNKeyNumLength

        SN4 = 4
        SN8 = 8
        SN12 = 12

    End Enum

    Public Class RandomSNKGenerator

        Private Shared Function AppendSpecifiedStr(ByVal length As Integer, ByVal str As String, ByVal newKey As Char())

            Dim newKeyStr As String = ""
            Dim k As Integer = 0

            For i As Integer = 0 To length - 1

                For k = i To 4 + (i - 1)
                    newKeyStr += newKey(k)
                Next

                If k = length Then
                    Exit For
                Else
                    i = k - 1
                    newKeyStr &= str
                End If

            Next

            Return newKeyStr
        End Function

        Public Shared Function GetSerialKeyAlphaNumaric(ByVal SNKeyLength As SNKeyLength, ByVal keyLength As Integer) As String

            Dim newguid As Guid = Guid.NewGuid()
            Dim randomStr As String = newguid.ToString("N")
            Dim tracStr As String = randomStr.Substring(0, keyLength)
            tracStr = tracStr.ToUpper()
            Dim newKey As Char() = tracStr.ToCharArray()
            Dim newSerialNumber As String = ""

            Select Case keyLength

                Case SNKeyLength.SN16
                    newSerialNumber = AppendSpecifiedStr(16, "-", newKey)

                Case SNKeyLength.SN20
                    newSerialNumber = AppendSpecifiedStr(20, "-", newKey)
                Case SNKeyLength.SN24
                    newSerialNumber = AppendSpecifiedStr(24, "-", newKey)
                Case SNKeyLength.SN28
                    newSerialNumber = AppendSpecifiedStr(28, "-", newKey)
                Case SNKeyLength.SN32
                    newSerialNumber = AppendSpecifiedStr(32, "-", newKey)
            End Select

            Return newSerialNumber

        End Function

        Public Shared Function GetSerialKeyNumaric(ByVal SNKeyNumLength As SNKeyNumLength, ByVal keyLength As Integer) As String

            Dim rn As New Random
            Dim sd As Double = Math.Round(rn.NextDouble() * Math.Pow(10, keyLength) + 4)
            Return sd.ToString().Substring(0, keyLength)
        End Function

    End Class

End Namespace