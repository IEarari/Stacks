Public Class Form1
    Dim Stack(3), Num, i, Sum, StackSize As Integer
    Dim Pointer As Integer = -1
    Dim inited As Boolean = False
    Public Sub init()
        If inited = False Then
            StackSize = InputBox("What is size of Stack do you want it to be?")
            Label16.Text = StackSize
            StackSize -= 1
            ReDim Stack(StackSize)
            inited = True
        Else
            MsgBox("You Can't Resize This Stack")
        End If
    End Sub
    Public Sub ADD()
        If Pointer = StackSize Then
            MsgBox("Stack Is Full")
        Else
            Label4.Text = Pointer
            Pointer += 1
            Label6.Text = Pointer
            Num = InputBox("Enter Item Value")
            Stack(Pointer) = Num
            ShowStackItems()
        End If
    End Sub

    Public Sub Remove()
        If Pointer = -1 Then
            MsgBox("Stack Is Empty")
        Else
            Label10.Text = Pointer
            Label14.Text = Stack(Pointer)
            Pointer -= 1
            Label12.Text = Pointer
            ShowStackItems()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Remove()
        ShowStackItemsSum()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If inited = True Then
            Label3.Text = ""
            Label4.Text = ""
            Label6.Text = ""
            Label8.Text = ""
            Label10.Text = ""
            Label12.Text = ""
            Label14.Text = ""
            Label16.Text = ""
            Pointer = -1
            inited = False
        Else
            MsgBox("There is No Data To Reset it")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        init()
    End Sub

    Public Sub ShowStackItems()
        Label3.Text = ""
        For i = 0 To Pointer Step 0.5
            Label3.Text = Label3.Text & " " & Stack(i)
            i += 1
        Next
    End Sub
    Public Sub ShowStackItemsSum()
        Sum = 0
        For i = 0 To Pointer Step 0.5
            Sum += Stack(i)
            i += 1
        Next
        Label8.Text = Sum
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If inited = True Then
            ADD()
            ShowStackItemsSum()
        Else
            MsgBox("Please Set Stack Size by clicking on start")
        End If
    End Sub

End Class