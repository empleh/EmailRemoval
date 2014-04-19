Public Class EmailRemoval

  Public Shared Function GenerateEmailReplacement(currentEmail As String,
                                                          emailBeingRemoved As String) As String
    Dim splitAddresses As List(Of String) = GetListOfAddresses(currentEmail)
    splitAddresses.RemoveAll(Function(address As String) AddressesMatch(address, emailBeingRemoved))

    Return String.Join(";", splitAddresses)
  End Function

  Public Shared Function GetListOfAddresses(ByVal emailAddresses As String) As List(Of String)
    Return emailAddresses.Split(";"c).ToList()
  End Function

  Public Shared Function AddressesMatch(toBeCompared As String,
                                        toBeRemoved As String) As Boolean
    Return (toBeCompared.Trim().ToUpper() = toBeRemoved.Trim().ToUpper())
  End Function

  'Public Shared Function GenerateEmailReplacement(currentEmail As String,
  '                                                emailBeingRemoved As String) As String
  '  Dim newEmail As String = currentEmail.Replace(emailBeingRemoved, "").Replace(";;", ";")
  '  If newEmail.EndsWith(";") Then
  '    newEmail = newEmail.Substring(0, newEmail.Length - 1)
  '  End If
  '  If newEmail.StartsWith(";") Then
  '    newEmail = newEmail.Substring(1)
  '  End If

  '  Return newEmail.Trim()
  'End Function
End Class

<TestClass()> Public Class GenerateEmailReplacementTests
    <TestMethod()>
    Public Sub WhenEmailMatchesExactlyReturnEmptyString()
        Const expected As String = ""

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("test@test.com", "test@test.com")

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod()>
    Public Sub WhenEmailDoesNotMatchReturnOriginalString()
        Const expected As String = "test@test.com"

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("test@test.com", "doesNotMatch@test.com")

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod()>
    Public Sub WhenMatchingAddressIsLastOfMultipleAddressesRemoveAddressAndLeadingSemiColon()
        Const expected As String = "test@test.com"

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("test@test.com;removeMe@test.com", "removeMe@test.com")

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod()>
    Public Sub WhenMatchingAddressIsInTheMiddleOfMultipleAddressesRemoveThatAddressAnd1SemiColon()
        Const expected As String = "test@test.com;second@test.com"

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("test@test.com;removeMe@test.com;second@test.com", "removeMe@test.com")

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod()>
    Public Sub WhenMatchingAddressIsTheFirstOfMultipleAddressesRemoveAddressAndTrailingSemiColon()
        Const expected As String = "test@test.com"

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("removeMe@test.com;test@test.com", "removeMe@test.com")

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod()>
    Public Sub WhenMatchingAddressBelongsToAPartOfAValidAddressTheValidAddressIsNotRemoved()
        Const expected As String = "ttest@test.com"

        Dim actual As String = EmailRemoval.GenerateEmailReplacement("ttest@test.com", "test@test.com")

        Assert.AreEqual(expected, actual)
    End Sub
End Class



