Class MainWindow
    Private counter As Integer = 1
    Private lettercards As IEnumerable(Of Object) = My.Settings.Letters.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
    Private misccards As IEnumerable(Of String) = My.Settings.Misc.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
    Private mixedcards As IEnumerable(Of String)
    Private tracker As IDictionary(Of String, Integer) = New Dictionary(Of String, Integer)

    Private Sub ChangeCard_Click(sender As Object, e As RoutedEventArgs) Handles ChangeCard.Click
        If (counter > mixedcards.Count()) Then
            MixCards()
            counter = 1
        End If
        Dim str As String = mixedcards(counter - 1)
        LetterBox.Text = str
        If (Not tracker.ContainsKey(str)) Then
            tracker.Add(str, 1)
        Else
            tracker(str) = tracker(str) + 1
        End If
        counter += 1

    End Sub

    Private Sub MixCards()
        mixedcards = lettercards.[Select](Function(x As String) x.ToUpper()) _
            .Concat(lettercards.[Select](Function(x As String) x.ToLower())) _
            .Concat(misccards) _
            .OrderBy(Function(a0 As String) Guid.NewGuid()).ToArray()
    End Sub

    Private Sub Window_Closed(sender As Object, e As EventArgs)

        Dim str As String = $".\{My.Settings.ProgressTrackerFile}"
        If (My.Computer.FileSystem.FileExists(str)) Then
            My.Computer.FileSystem.DeleteFile(str)
        End If
        Dim stringBuilder As Text.StringBuilder = New Text.StringBuilder()
        For Each t In tracker.OrderBy(Function(z) z.Key)
            stringBuilder.AppendLine(String.Format("{0} , {1}", t.Key, t.Value))
        Next
        My.Computer.FileSystem.WriteAllText(str, stringBuilder.ToString(), True)

    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        MixCards()
    End Sub
End Class
