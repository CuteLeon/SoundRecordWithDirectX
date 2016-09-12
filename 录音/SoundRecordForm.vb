Public Class SoundRecordForm
    '说明：需要在 [App.config] 文件里把 "<startup>" 改成 "<startup useLegacyV2RuntimeActivationPolicy="true">"

    Dim SoundRecorder As SoundRecord
    Dim FileName As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SoundRecorder = New SoundRecord
        FileName = IO.Path.GetTempPath & "Sound-" & My.Computer.Clock.TickCount.ToString & ".wav"
        SoundRecorder.SetFileName(FileName)
        SoundRecorder.RecordStart()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SoundRecorder.RecordStop()
        SoundRecorder = Nothing
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            My.Computer.Audio.Play(FileName)
            Kill(FileName)
            Me.Text = "播放完毕：" & My.Computer.Clock.LocalTime.ToLongTimeString
        Catch ex As Exception
            Me.Text = "播放出错：" & My.Computer.Clock.LocalTime.ToLongTimeString
        End Try
    End Sub
End Class
