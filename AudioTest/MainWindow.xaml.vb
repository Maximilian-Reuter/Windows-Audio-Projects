Imports AudioTest

Class MainWindow
    Implements Interface1

    Public Property test

    Private Sub abcdefg(sender As Object, e As RoutedEventArgs) Handles button.Click

    End Sub

    Private Sub button_Copy_Click(sender As Object, e As RoutedEventArgs) Handles button_Copy.Click

    End Sub

    Public Function tschüss(S As String) As String Implements Interface1.hallo
        Throw New NotImplementedException()
    End Function
End Class
