Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Text
Imports System.Windows.Forms
Imports System.Diagnostics
Public Class Form1
    Inherits Form
    Implements IATMActions

#Region "Global variables"
    ' Credentials for email.
    Private Const EUser As String = "sarahi.reyessrv@outlook.com"
    Private Const Password As String = "UTAU-0401"
    Private Enum ATMstatus
        Home
        Deposit
        Withdrawal
        W1
        W2
        W3
        W4
        Query
        AddBills
    End Enum
    Private Actualstatus As ATMstatus
    Public ReadOnly database As String = "server=HP-SARAHI\SARAHIDB; database=atmDB; integrated security=true"
    Private BalanceActualUser As Integer
    Private NameActualUser As String
    Private txtamount As Integer
    Private totalAmount As Integer
    Private timer As Timer
    Private user As User
    Private bancknote100 As Integer
    Private bancknote200 As Integer
    Private bancknote500 As Integer
    Private newBalance As Integer
    Private connection As SqlConnection
#End Region
    Public Sub New()
        InitializeComponent()
        WelcomePanel.Visible = True
        OptionsPanel.Visible = False
        DepositPanel.Visible = False
        SelectedBillPanel.Visible = False
        NotFoundsPanel.Visible = False
        WithdrawalPanel.Visible = False
        withdrawalSuccessfullPanel.Visible = False
        ErrorAmountPanel.Visible = False
        ErrorBillsPanel.Visible = False
        ErrorHomePanel.Visible = False
        ErrorDispensedPanel.Visible = False
        QueryPanel.Visible = False
        insufficientBalancePanel.Visible = False
        timer = New Timer()
        timer.Interval = 5000 ' 5 seconds
        AddHandler timer.Tick, AddressOf BtnCancel_Click
        connection = New SqlConnection(database)
    End Sub

#Region "Method´s IATMActions"
    Public Sub Deposit(amount As Integer) Implements IATMActions.Deposit
        bancknote100 = CInt(nud100.Value)
        bancknote200 = CInt(nud200.Value)
        bancknote500 = CInt(nud500.Value)
        totalAmount = (bancknote100 * 100) + (bancknote200 * 200) + (bancknote500 * 500)

        If totalAmount = txtamount Then
            Using connection As New SqlConnection(database)
                connection.Open()

                Dim updateQuery As String = "UPDATE [Money] SET Quantity = Quantity + @Quantity WHERE Denomination = @Denomination"

                Using command As New SqlCommand(updateQuery, connection)
                    command.Parameters.AddWithValue("@Quantity", bancknote100)
                    command.Parameters.AddWithValue("@Denomination", 100)
                    command.ExecuteNonQuery()

                    command.Parameters.Clear()
                    command.Parameters.AddWithValue("@Quantity", bancknote200)
                    command.Parameters.AddWithValue("@Denomination", 200)
                    command.ExecuteNonQuery()

                    command.Parameters.Clear()
                    command.Parameters.AddWithValue("@Quantity", bancknote500)
                    command.Parameters.AddWithValue("@Denomination", 500)
                    command.ExecuteNonQuery()
                End Using
            End Using

            newBalance = BalanceActualUser + txtamount
            UpdateBalanceInDatabase(newBalance)
            UpdateUserBalance()
            'UpdateBalanceInDatabase(newBalance)
            Console.Beep()
            DepositSuccessfullLabel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()
        Else
            Console.Beep()
            SelectedBillPanel.Visible = False
            ErrorBillsPanel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()
            Return
        End If
    End Sub
    Public Sub Withdrawal(WAmount As Integer) Implements IATMActions.Withdrawal
        If WAmount <= 0 Then
            Console.Beep()
            WithdrawalPanel.Visible = False
            ErrorAmountPanel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()
            Return
        End If

        If WAmount > BalanceActualUser Then
            Console.Beep()
            WithdrawalPanel.Visible = False
            insufficientBalancePanel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()
            Return
        End If

        If Not HasAnyFunds() Then
            Console.Beep()
            WithdrawalPanel.Visible = False
            NotFoundsPanel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()

            Dim De As String = "sarahi.reyessrv@outlook.com"
            Dim Para As String = "sarahi.reyessrv@outlook.com"
            Dim Asunto As String = "Insufficient funds alert at the ATM"
            Dim Mensaje As String = "It has been confirmed that the ATM does not have available bills for withdrawals." & vbCrLf & "Please request an administrator to refill the bills in the ATM's database." & vbCrLf & "Thank you."
            Dim ErrorEmail As String = ""
            Dim MensajeBuiler As New StringBuilder()
            MensajeBuiler.Append(Mensaje)
            SendEmail(MensajeBuiler, DateTime.Now, De, Para, Asunto, ErrorEmail)
            Return
        End If

        If Not CanDispenseAmount(WAmount) Then
            Console.Beep()
            WithdrawalPanel.Visible = False
            ErrorDispensedPanel.Visible = True
            BtnEnter.Enabled = False
            timer.Start()
            Return
        End If

        DispenseAmount(WAmount)
        newBalance = BalanceActualUser - WAmount
        UpdateBalanceInDatabase(newBalance)
        UpdateUserBalance()
        CreateTransactionFile(DateTime.Now, newBalance, WAmount)

        Console.Beep()
        WithdrawalPanel.Visible = False
        withdrawalSuccessfullPanel.Visible = True
        BtnEnter.Enabled = False
        timer.Start()
    End Sub
#End Region

#Region "Method´s"
    Private Function HasAnyFunds() As Boolean
        Dim denominations As Integer() = New Integer() {500, 200, 100}

        For Each denomination As Integer In denominations
            Dim availableQuantity As Integer = GetAvailableQuantity(denomination)

            If availableQuantity > 0 Then
                Return True
            End If
        Next

        Return False
    End Function
    Private Sub UpdateBalanceInDatabase(newBalance As Integer)
        Using connection As New SqlConnection(database)
            Try
                connection.Open()
            Catch ex As Exception
                Console.WriteLine("Error al conectar a la base de datos.")
                Return
            End Try

            Dim query As String = "UPDATE [user] SET Balance = @Balance WHERE id = @Id AND name = @Name"
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Balance", newBalance)
                command.Parameters.AddWithValue("@Id", user.ID)
                command.Parameters.AddWithValue("@Name", NameActualUser)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Function GetAvailableQuantity(denomination As Integer) As Integer
        Dim quantity As Integer = 0

        Using connection As New SqlConnection(database)
            connection.Open()

            Dim query As String = "SELECT Quantity FROM [Money] WHERE Denomination = @Denomination"
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Denomination", denomination)
                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    quantity = DirectCast(result, Integer)
                End If
            End Using
        End Using

        Return quantity
    End Function
    Private Function CanDispenseAmount(withdrawalAmount As Integer) As Boolean
        Dim remainingAmount As Integer = withdrawalAmount
        Dim denominations As Integer() = New Integer() {500, 200, 100}

        For Each denomination As Integer In denominations
            Dim availableQuantity As Integer = GetAvailableQuantity(denomination)
            Dim requiredQuantity As Integer = remainingAmount \ denomination

            If requiredQuantity > availableQuantity Then
                remainingAmount -= availableQuantity * denomination
            Else
                remainingAmount = remainingAmount Mod denomination
            End If
        Next

        Return remainingAmount = 0
    End Function

    Private Sub DispenseAmount(withdrawalAmount As Integer)
        Dim remainingAmount As Integer = withdrawalAmount
        Dim denominations As Integer() = New Integer() {500, 200, 100}

        Using connection As New SqlConnection(database)
            connection.Open()

            For Each denomination As Integer In denominations
                Dim availableQuantity As Integer = GetAvailableQuantity(denomination)
                Dim requiredQuantity As Integer = remainingAmount \ denomination
                Dim dispensedQuantity As Integer = Math.Min(requiredQuantity, availableQuantity)

                If dispensedQuantity > 0 Then
                    Dim updateQuery As String = "UPDATE [Money] SET Quantity = Quantity - @Quantity WHERE Denomination = @Denomination"
                    Using command As New SqlCommand(updateQuery, connection)
                        command.Parameters.AddWithValue("@Quantity", dispensedQuantity)
                        command.Parameters.AddWithValue("@Denomination", denomination)
                        command.ExecuteNonQuery()
                    End Using

                    remainingAmount -= dispensedQuantity * denomination
                End If
            Next
        End Using
    End Sub

    Private Function GetBalanceFromDatabase() As Integer
        Dim balance As Integer = 0

        Using connection As New SqlConnection(database)
            connection.Open()

            Dim query As String = "SELECT Balance FROM [user] WHERE id = @Id AND name = @Name"
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Id", user.ID)
                command.Parameters.AddWithValue("@Name", NameActualUser)

                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    balance = CInt(result)
                End If
            End Using
        End Using

        Return balance
    End Function

    Private Sub UpdateUserBalance()
        BalanceActualUser = GetBalanceFromDatabase()
    End Sub
    Private Sub CreateTransactionFile(dateTime As DateTime, currentBalance As Integer, withdrawalAmount As Integer)
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim fileName As String = "Transaction" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        Dim filePath As String = Path.Combine(desktopPath, fileName)

        Using writer As New StreamWriter(filePath)
            writer.WriteLine("User: " & NameActualUser.ToUpper())
            writer.WriteLine("Date: " & dateTime.ToString("yyyy-MM-dd"))
            writer.WriteLine("Time: " & dateTime.ToString("hh:mm:ss tt"))
            writer.WriteLine("Withdrawal transaction")
            writer.WriteLine("Withdrawal Amount: " & String.Format("{0:C2}", withdrawalAmount) & " MXN")
            writer.WriteLine("Dispensed Bills:")

            Dim denominations As Integer() = New Integer() {500, 200, 100}
            For Each denomination As Integer In denominations
                Dim dispensedQuantity As Integer = withdrawalAmount \ denomination
                If dispensedQuantity > 0 Then
                    writer.WriteLine(dispensedQuantity & " x " & String.Format("{0:C2}", denomination) & " MXN")
                    withdrawalAmount = withdrawalAmount Mod denomination
                End If
            Next

            writer.WriteLine("Current balance: " & String.Format("{0:C2}", currentBalance) & " MXN")
            writer.WriteLine("Thank you for your preference!")
        End Using
    End Sub
    Public Shared Sub SendEmail(ByRef Mensaje As StringBuilder, FechaEnvio As DateTime, De As String, Para As String, Asunto As String, ByRef [Error] As String)
        [Error] = ""
        Try
            Mensaje.Append(Environment.NewLine)
            Mensaje.Append(String.Format("This email was sent on {0:dd/MM/yyyy} at {0:HH:mm:ss} Hrs: " & vbCr & vbLf, FechaEnvio))
            Mensaje.Append(Environment.NewLine)
            Dim mail As New MailMessage()
            mail.From = New MailAddress(De)
            mail.To.Add(Para)
            mail.Subject = Asunto
            mail.Body = Mensaje.ToString()
            Dim smtp As New SmtpClient("smtp.office365.com")
            smtp.Port = 587
            smtp.UseDefaultCredentials = False
            smtp.Credentials = New System.Net.NetworkCredential(EUser, Password)
            smtp.EnableSsl = True
            smtp.Send(mail)
            [Error] = "Email successfully sent"
        Catch ex As Exception
            [Error] = "Error: " & ex.Message
            MessageBox.Show([Error])
            Return
        End Try
    End Sub
#End Region

#Region "Others"
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        timer.Stop()
    End Sub

    Protected Overrides Sub Finalize()
        timer.Stop()
        timer.Dispose()

        If connection IsNot Nothing AndAlso connection.State = ConnectionState.Open Then
            connection.Close()
            connection.Dispose()
        End If

        MyBase.Finalize()
    End Sub

    Private Sub AuthorLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles AuthorLabel.LinkClicked
        Dim psi As New ProcessStartInfo("https://github.com/Sarahi-I22050355/ATMAppVB") With {.UseShellExecute = True}

        Process.Start(psi)
    End Sub

    Private Sub BtnNumber_Click(sender As Object, e As EventArgs) Handles Btn9.Click, Btn8.Click, Btn7.Click, Btn6.Click, Btn5.Click, Btn4.Click, Btn3.Click, Btn0.Click, Btn00.Click, Btn2.Click, Btn1.Click
        Dim button As Button = DirectCast(sender, Button)
        Dim number As String = button.Text
        If DepositPanel.Visible OrElse WithdrawalPanel.Visible Then
            txtBoxDepositAmount.Text += number
            txtBoxWithdrawalAmount.Text += number
        Else
            Return
        End If
    End Sub
#End Region

#Region "Clear - Cancel - Enter"
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click, timer1.Tick
        Actualstatus = ATMstatus.Home
        UpdateUserBalance()
        OptionsPanel.Visible = True
        ErrorAmountPanel.Visible = False
        ErrorHomePanel.Visible = False
        ErrorBillsPanel.Visible = False
        ErrorDispensedPanel.Visible = False
        DepositPanel.Visible = False
        QueryPanel.Visible = False
        insufficientBalancePanel.Visible = False
        withdrawalSuccessfullPanel.Visible = False
        WithdrawalPanel.Visible = False
        SelectedBillPanel.Visible = False
        NotFoundsPanel.Visible = False
        txtBoxDepositAmount.Text = String.Empty
        txtBoxWithdrawalAmount.Text = String.Empty
        BtnEnter.Enabled = True
        txtamount = 0
        bancknote100 = 0
        bancknote200 = 0
        bancknote500 = 0
        totalAmount = 0
        newBalance = 0
        nud100.Value = 0
        nud200.Value = 0
        nud500.Value = 0
        DepositSuccessfullLabel.Visible = False
        Timer1_Tick(sender, e)
        Console.Beep()
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        If DepositPanel.Visible Then
            If Not String.IsNullOrEmpty(txtBoxDepositAmount.Text) Then
                txtBoxDepositAmount.Text = txtBoxDepositAmount.Text.Remove(txtBoxDepositAmount.Text.Length - 1)
            End If
        ElseIf WithdrawalPanel.Visible Then
            If Not String.IsNullOrEmpty(txtBoxWithdrawalAmount.Text) Then
                txtBoxWithdrawalAmount.Text = txtBoxWithdrawalAmount.Text.Remove(txtBoxWithdrawalAmount.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub BtnEnter_Click(sender As Object, e As EventArgs) Handles BtnEnter.Click
        Select Case Actualstatus
            Case ATMstatus.Deposit
                If Not DepositPanel.Visible Then
                    Return
                Else
                    Dim amount As Integer
                    If Integer.TryParse(txtBoxDepositAmount.Text, amount) AndAlso amount Mod 100 = 0 AndAlso amount > 99 Then
                        DepositPanel.Visible = False
                        Console.Beep()
                        SelectedBillPanel.Visible = True
                        Actualstatus = ATMstatus.AddBills
                        txtamount = amount
                    Else
                        DepositPanel.Visible = False
                        ErrorAmountPanel.Visible = True
                        BtnEnter.Enabled = False
                        timer.Start()
                        Return
                    End If
                End If
            Case ATMstatus.AddBills
                DirectCast(Me, IATMActions).Deposit(txtamount)
            Case ATMstatus.Withdrawal
                If Not WithdrawalPanel.Visible Then
                    Return
                Else
                    Dim WAmount As Integer
                    If Integer.TryParse(txtBoxWithdrawalAmount.Text, WAmount) AndAlso WAmount Mod 100 = 0 AndAlso WAmount > 99 Then
                        WithdrawalPanel.Visible = False
                        Console.Beep()
                        BtnEnter.Enabled = False
                        txtamount = WAmount
                        DirectCast(Me, IATMActions).Withdrawal(txtamount)
                    Else
                        WithdrawalPanel.Visible = False
                        ErrorAmountPanel.Visible = True
                        BtnEnter.Enabled = False
                        timer.Start()
                        Return
                    End If
                End If
            Case ATMstatus.Query
                BtnEnter.Enabled = False
                BtnCancel_Click(sender, e)
            Case ATMstatus.Home
                ErrorHomePanel.Visible = True
                UpdateUserBalance()
                Console.Beep()
                BtnEnter.Enabled = False
                timer.Start()
            Case Else
        End Select
    End Sub
#End Region

#Region "ATM Oprtions"
    Private Sub BtnDeposit_Click(sender As Object, e As EventArgs) Handles BtnDeposit.Click
        OptionsPanel.Visible = False
        Console.Beep()
        SelectedBillPanel.Visible = False
        DepositPanel.Visible = True
        Actualstatus = ATMstatus.Deposit
    End Sub

    Private Sub BtnWithdrawal_Click(sender As Object, e As EventArgs) Handles BtnWithdrawal.Click
        OptionsPanel.Visible = False
        DepositPanel.Visible = False
        SelectedBillPanel.Visible = False
        Console.Beep()
        withdrawalSuccessfullPanel.Visible = False
        WithdrawalPanel.Visible = True
        Actualstatus = ATMstatus.Withdrawal
    End Sub

    Private Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles BtnQuery.Click
        Console.Beep()
        Actualstatus = ATMstatus.Query
        Dim currentBalance As Integer = GetBalanceFromDatabase()
        QueryPanel.Visible = True
        QueryLabel.Text = (String.Format("{0:C2}", currentBalance) & " MXN")
    End Sub

    Private Sub BtnLogOut_Click(sender As Object, e As EventArgs) Handles BtnLogOut.Click
        user.NIP = Nothing
        user.ID = 0
        user.Name = Nothing
        user.Balance = 0
        textBoxID.Clear()
        textBoxNIP.Clear()
        textBoxN.Clear()
        Console.Beep()
        WelcomePanel.Visible = True
    End Sub

    Private Sub BtnVerify_Click(sender As Object, e As EventArgs) Handles BtnVerify.Click
        user = New User()
        Dim _id As Integer = Nothing

        If String.IsNullOrWhiteSpace(textBoxID.Text) OrElse String.IsNullOrWhiteSpace(textBoxN.Text) OrElse String.IsNullOrWhiteSpace(textBoxNIP.Text) Then
            MessageBox.Show("Please enter ID, name, and NIP.")
            textBoxID.Clear()
            textBoxNIP.Clear()
            textBoxN.Clear()
            Return
        Else

            Using connection As SqlConnection = New SqlConnection(database)

                Try
                    connection.Open()
                Catch __unusedException1__ As Exception
                    MessageBox.Show("Error connecting to the database.")
                    Return
                End Try

                If Not Integer.TryParse(textBoxID.Text, _id) Then
                    MessageBox.Show("Invalid ID. Please enter a numeric value.")
                    textBoxID.Clear()
                    textBoxNIP.Clear()
                    textBoxN.Clear()
                    Return
                End If

                user.ID = _id
                Dim _name As String = textBoxN.Text
                user.Name = _name
                Dim query As String = "SELECT * FROM [user] WHERE id = @Id AND name = @Name"

                Using command As SqlCommand = New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@Id", _id)
                    command.Parameters.AddWithValue("@Name", _name)

                    Using reader As SqlDataReader = command.ExecuteReader()

                        If Not reader.Read() Then
                            MessageBox.Show("Invalid credentials. Please try again.")
                            textBoxID.Clear()
                            textBoxNIP.Clear()
                            textBoxN.Clear()
                            Return
                        Else
                            Dim storedNIP As String = reader("NIP").ToString()

                            If storedNIP <> textBoxNIP.Text Then
                                MessageBox.Show("Invalid credentials. Please try again.")
                                textBoxID.Clear()
                                textBoxNIP.Clear()
                                textBoxN.Clear()
                                Return
                            End If

                            BalanceActualUser = Convert.ToInt32(reader("Balance"))
                            NameActualUser = reader("Name").ToString()
                            Console.Beep()
                            WelcomePanel.Visible = False
                            OptionsPanel.Visible = True
                            UserLabel.Text = user.Name.ToUpper()
                            Actualstatus = ATMstatus.Home
                        End If
                    End Using
                End Using
            End Using
        End If

    End Sub
#End Region

End Class
