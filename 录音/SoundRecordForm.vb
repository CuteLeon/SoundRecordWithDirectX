Public Class SoundRecordForm
    '说明：需要在 [App.config] 文件里把 "<startup>" 改成 "<startup useLegacyV2RuntimeActivationPolicy="true">"

    Dim SoundRecorder As SoundRecord

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SoundRecorder = New SoundRecord
        Dim wavfile As String = "test_" & My.Computer.Clock.TickCount & ".wav"
        SoundRecorder.SetFileName(wavfile)
        SoundRecorder.RecStart()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SoundRecorder.RecStop()
        SoundRecorder = Nothing
    End Sub
End Class
