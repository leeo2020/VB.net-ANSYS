Imports System
Public Class Form1
    Dim foldername As String
    Dim filename As String
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Button1.Font = New Font("新宋体", 14, FontStyle.Bold)
        Button1.AutoSize = True
        Button1.Size = New Size(200, 50)
        Button1.Text = "调用ANSYS后台计算"
        Me.Text = "ANSYS计算助手"
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim X
        Dim workdir As String
        Dim ansysexe As String
        Dim inputfile As String
        Dim outfile As String
        ansysexe = "E:\Technology\Mechanical\FEA\Ansys15.0\ANSYS Inc\v150\ansys\bin\winx64\ANSYS150.exe"
        workdir = folderchoice()
        inputfile = filechoice(workdir)
        outfile = "\result.log"

        If Dir(workdir & "\file.err") <> "" Then
            Kill(workdir & "\file.err")
        End If
        If Dir(inputfile) <> "" Then
            X = Shell(ansysexe & Chr(32) & "-b -p ane3fl -dir" & Chr(32) & workdir & Chr(32) & "-l en-us -s read -i" & Chr(32) & inputfile & Chr(32) & "-o" & Chr(32) & workdir & outfile, vbMinimizedFocus)
            'X = Shell(ansysexe & Chr(32) & "-b -p ane3fl -dir" & Chr(32) & workdir & Chr(32) & "-l en-us -s read -i" & Chr(32) & workdir & inputfile & Chr(32) & "-o" & Chr(32) & workdir & outfile, vbMinimizedFocus)
            'X = Shell("E:\Technology\Mechanical\FEA\Ansys15.0\ANSYS Inc\v150\ansys\bin\winx64\ANSYS150.exe -b -p ane3fl -dir F:\CGN\DailyTask\Windduct -l en-us -s read -i F:\CGN\DailyTask\Windduct\mymodel.txt -o F:\CGN\DailyTask\Windduct\jieguo.log", vbMinimizedFocus)
        End If
    End Sub
    ''选择文件
    Function filechoice(workdir As String) As String
        Dim ffilename As String
        ffilename = ""
        OpenFileDialog1.Title = "请选择计算文件"
        OpenFileDialog1.InitialDirectory = workdir
        OpenFileDialog1.DefaultExt = ".txt"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ffilename = OpenFileDialog1.FileName
        End If
        Return ffilename
    End Function
    ''选择工作文件
    Function folderchoice() As String
        Dim ffoldername As String
        ffoldername = ""
        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ffoldername = FolderBrowserDialog1.SelectedPath
        End If
        Return ffoldername
    End Function
End Class
