Imports System.Threading
Imports System.IO

Public Class frmMain
    Inherits System.Windows.Forms.Form
    Dim dontsave As Boolean = True
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents nudDiscount As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudLess As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents nudLowSale As System.Windows.Forms.NumericUpDown
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dtpstart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpend As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents F As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents txtmni As System.Windows.Forms.TextBox
    Friend WithEvents txtqb As System.Windows.Forms.TextBox
    Friend WithEvents O As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button4 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dtpend = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpstart = New System.Windows.Forms.DateTimePicker
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.txtqb = New System.Windows.Forms.TextBox
        Me.txtmni = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.nudLowSale = New System.Windows.Forms.NumericUpDown
        Me.Label11 = New System.Windows.Forms.Label
        Me.nudLess = New System.Windows.Forms.NumericUpDown
        Me.nudDiscount = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.F = New System.Windows.Forms.FolderBrowserDialog
        Me.O = New System.Windows.Forms.OpenFileDialog
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.nudLowSale, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLess, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(280, 262)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dtpend)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.dtpstart)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(272, 236)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Start"
        '
        'dtpend
        '
        Me.dtpend.Location = New System.Drawing.Point(42, 110)
        Me.dtpend.Name = "dtpend"
        Me.dtpend.Size = New System.Drawing.Size(200, 21)
        Me.dtpend.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Date end:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Date start:"
        '
        'dtpstart
        '
        Me.dtpstart.Location = New System.Drawing.Point(40, 56)
        Me.dtpstart.Name = "dtpstart"
        Me.dtpstart.Size = New System.Drawing.Size(200, 21)
        Me.dtpstart.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(94, 155)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Go!"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Button4)
        Me.TabPage2.Controls.Add(Me.Button3)
        Me.TabPage2.Controls.Add(Me.Button2)
        Me.TabPage2.Controls.Add(Me.txtqb)
        Me.TabPage2.Controls.Add(Me.txtmni)
        Me.TabPage2.Controls.Add(Me.Label13)
        Me.TabPage2.Controls.Add(Me.Label12)
        Me.TabPage2.Controls.Add(Me.nudLowSale)
        Me.TabPage2.Controls.Add(Me.Label11)
        Me.TabPage2.Controls.Add(Me.nudLess)
        Me.TabPage2.Controls.Add(Me.nudDiscount)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(272, 236)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Option"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(56, 200)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(16, 21)
        Me.Button4.TabIndex = 24
        Me.Button4.Text = "!"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(232, 200)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(24, 21)
        Me.Button3.TabIndex = 23
        Me.Button3.Text = "..."
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(232, 160)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(24, 21)
        Me.Button2.TabIndex = 22
        Me.Button2.Text = "..."
        '
        'txtqb
        '
        Me.txtqb.Location = New System.Drawing.Point(72, 200)
        Me.txtqb.Name = "txtqb"
        Me.txtqb.ReadOnly = True
        Me.txtqb.Size = New System.Drawing.Size(160, 21)
        Me.txtqb.TabIndex = 21
        '
        'txtmni
        '
        Me.txtmni.Location = New System.Drawing.Point(72, 160)
        Me.txtmni.Name = "txtmni"
        Me.txtmni.ReadOnly = True
        Me.txtmni.Size = New System.Drawing.Size(160, 21)
        Me.txtmni.TabIndex = 20
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 184)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(144, 16)
        Me.Label13.TabIndex = 19
        Me.Label13.Text = "Quickbooks path:"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(24, 144)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(144, 16)
        Me.Label12.TabIndex = 18
        Me.Label12.Text = "MNI path:"
        '
        'nudLowSale
        '
        Me.nudLowSale.Location = New System.Drawing.Point(72, 120)
        Me.nudLowSale.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nudLowSale.Name = "nudLowSale"
        Me.nudLowSale.Size = New System.Drawing.Size(72, 21)
        Me.nudLowSale.TabIndex = 17
        Me.nudLowSale.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudLowSale.ThousandsSeparator = True
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(24, 104)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(144, 16)
        Me.Label11.TabIndex = 16
        Me.Label11.Text = "Low Sales:"
        '
        'nudLess
        '
        Me.nudLess.Location = New System.Drawing.Point(72, 80)
        Me.nudLess.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nudLess.Name = "nudLess"
        Me.nudLess.Size = New System.Drawing.Size(72, 21)
        Me.nudLess.TabIndex = 3
        Me.nudLess.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudLess.ThousandsSeparator = True
        '
        'nudDiscount
        '
        Me.nudDiscount.Location = New System.Drawing.Point(72, 32)
        Me.nudDiscount.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nudDiscount.Name = "nudDiscount"
        Me.nudDiscount.Size = New System.Drawing.Size(72, 21)
        Me.nudDiscount.TabIndex = 2
        Me.nudDiscount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudDiscount.ThousandsSeparator = True
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Maximum Less:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Maximum Discount:"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(280, 262)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.nudLowSale, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLess, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public del As New ListBox
    Dim formprog As New frmProg
    Public formreport As New frmReport
    Public formmain As frmMain



    Private Sub ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudDiscount.ValueChanged, nudLess.ValueChanged, nudLowSale.ValueChanged, txtmni.TextChanged, txtqb.TextChanged
        save()
    End Sub

    Public Sub save()
        If dontsave = True Then Exit Sub
        FileOpen(1, Application.StartupPath & "\data.ini", OpenMode.Output, OpenAccess.Write)
        Write(1, Me.dtpstart.Value, Me.dtpend.Value, Me.nudDiscount.Value, Me.nudLess.Value, Me.nudLowSale.Value, Me.txtmni.Text, Me.txtqb.Text)
        FileClose(1)
    End Sub

    Public Sub loads()
        On Error Resume Next
        Dim date1 As Date
        Dim date2 As Date
        Dim dec1 As Decimal
        Dim dec2 As Decimal
        Dim dec3 As Decimal
        Dim str1 As String = ""
        Dim str2 As String = ""

        FileOpen(1, Application.StartupPath & "\Data.ini", OpenMode.Input, OpenAccess.Read)
        Input(1, date1)
        Input(1, date2)

        Input(1, dec1)
        Input(1, dec2)
        Input(1, dec3)

        Input(1, str1)
        Input(1, str2)
        FileClose(1)

        Me.dtpstart.Value = date1
        Me.dtpend.Value = date2
        Me.nudDiscount.Value = dec1
        Me.nudLess.Value = dec2

        Me.nudLowSale.Value = dec3
        Me.txtmni.Text = str1
        Me.txtqb.Text = str2
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtmni.Text = Application.StartupPath & "\MNI"
        txtqb.Text = Application.StartupPath & "\EXCEL"
        loads()
        dontsave = False
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        report()

    End Sub
    Public Sub report()
        'On Error Resuformmain Next
        formprog.Hide()
        formreport.Show()
        Dim i As Integer

        Me.Hide()

        formprog.Show()

        'creating lstview
        formprog.Label1.Text = "creating report"

        For i = 0 To DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)
            Application.DoEvents()
            formreport.lsvSales.BeginUpdate()
            If DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value) <= 0 Then
                MsgBox("Pls check your date stupid! Stopping...")
                End
            End If
            formprog.ProgressBar1.Value = Int((i / DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)) * 100 / 4)
            formprog.Text = "progress - " & Int((i / DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)) * 100 / 4) & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()
            Dim litem As New ListViewItem
            If DateAdd(DateInterval.Day, i, dtpstart.Value).DayOfWeek = DayOfWeek.Sunday Then litem.ForeColor = Color.Red
            litem.Text = Format(DateAdd(DateInterval.Day, i, dtpstart.Value), "long date")
            litem.Tag = Format(DateAdd(DateInterval.Day, i, dtpstart.Value), "short date")
            litem.SubItems.Add("-----")
            litem.SubItems.Add("-----")
            litem.SubItems.Add("-----")
            litem.SubItems.Add("-----")
            formreport.lsvSales.Items.Add(litem)
            formreport.lsvSales.EndUpdate()
        Next
        'opening mni
        formprog.Label1.Text = "calculating *.mni"
        For i = 0 To DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)
            Application.DoEvents()
            formprog.ProgressBar1.Value = Int((i / DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)) * 100 / 4) + 25
            formprog.Text = "progress - " & Int((i / DateDiff(DateInterval.Day, dtpstart.Value, dtpend.Value)) * 100 / 4) + 25 & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()
            Dim less As Double
            Dim discount As Double
            Dim sales As Double
            If MNI(txtmni.Text & "/" & Format(DateAdd(DateInterval.Day, i, dtpstart.Value), "MM_dd_yy") & ".mni", less, discount, sales) = False Then
            Else
                formreport.lsvSales.Items(i).SubItems(1).Text = Format(less, "#,##0.00")
                formreport.lsvSales.Items(i).SubItems(2).Text = Format(discount, "#,##0.00")
                formreport.lsvSales.Items(i).SubItems(3).Text = Format(sales, "#,##0.00")
            End If
        Next
        'opening qb
        formprog.Label1.Text = "reading quickbooks file"
        Application.DoEvents()
        formprog.Refresh()
        formprog.ProgressBar1.Refresh()

        Dim sessManager As New QBSessionManager
        Dim msgSetRs As IMsgSetResponse


        Dim authFlags As Long
        authFlags = 0
        authFlags = authFlags Or &H1&
        Dim AuthPrefs As IQBAuthPreferences
        AuthPrefs = sessManager.QBAuthPreferences
        AuthPrefs.PutAuthFlags(authFlags)

        Try
            Application.DoEvents()
            Dim msgSetRq As IMsgSetRequest = sessManager.CreateMsgSetRequest("US", 6, 0)
            'msgSetRq.ClearRequests()
            msgSetRq.Attributes.OnError = ENRqOnError.roeContinue

            Application.DoEvents()
            formprog.ProgressBar1.Value = 51
            formprog.Text = "progress - " & formprog.ProgressBar1.Value & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()


            Dim custAdd As ISalesReceiptQuery = msgSetRq.AppendSalesReceiptQueryRq

            ' to fasten d sync of data
            custAdd.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(dtpstart.Value)
            custAdd.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(dtpend.Value)
            custAdd.IncludeLineItems.SetValue(True)


            Application.DoEvents()
            formprog.ProgressBar1.Value = 52
            formprog.Text = "progress - " & formprog.ProgressBar1.Value & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()


            'step3: begin QB session and send the request
            sessManager.OpenConnection("Dona", "Dona of BOX")
            Application.DoEvents()
            formprog.ProgressBar1.Value = 53
            formprog.Text = "progress - " & formprog.ProgressBar1.Value & "% - transferring data..."
            formprog.Label1.Text = "reading quickbooks file... transfering data... wait a momment..."

            formprog.Refresh()
            formprog.ProgressBar1.Refresh()

            sessManager.BeginSession(txtqb.Text, ENOpenMode.omDontCare)
            Application.DoEvents()
            msgSetRs = sessManager.DoRequests(msgSetRq)
            sessManager.EndSession()
            sessManager.CloseConnection()
            'frm2.TextBox1.Text = msgSetRs.ToXMLString
            'frm2.Show()


            Application.DoEvents()
            formprog.ProgressBar1.Value = 54
            formprog.Text = "progress - " & formprog.ProgressBar1.Value & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()

            Dim SRList As ISalesReceiptRetList = msgSetRs.ResponseList.GetAt(0).Detail
            Dim looper As Integer
            For looper = 0 To formreport.lsvSales.Items.Count - 1
                Application.DoEvents()
                formprog.ProgressBar1.Value = Int((looper / (formreport.lsvSales.Items.Count - 1)) * 100 / 5) + 55
                formprog.Text = "progress - " & formprog.ProgressBar1.Value & "%"
                formprog.Refresh()
                formprog.ProgressBar1.Refresh()

                Dim looper1 As Integer
                'My brain is upside down

                For looper1 = SRList.Count - 1 To 0 Step -1
                    Application.DoEvents()

                    If CDate(formreport.lsvSales.Items(looper).Text) = SRList.GetAt(looper1).DueDate.GetValue Then
                        Dim Cusage As Double = 0
                        Dim BW As Double = 0
                        Dim Colored As Double = 0
                        Dim Scan As Double = 0

                        Dim looper2 As Integer

                        For looper2 = SRList.GetAt(looper1).ORSalesReceiptLineRetList.Count - 1 To 0 Step -1
                            Dim srl As ISalesReceiptLineRet = SRList.GetAt(looper1).ORSalesReceiptLineRetList.GetAt(looper2).SalesReceiptLineRet
                            Application.DoEvents()
                            Select Case srl.ItemRef.FullName.GetValue 'Str(SRList.GetAt(looper1).ORSalesReceiptLineRetList.GetAt(looper2).SalesReceiptLineRet.ItemRef.FullName.GetValue)
                                Case "Computer Income:Usage"
                                    Cusage = Val(srl.Amount.GetValue.ToString)
                                Case "Computer Income:BW Printing"
                                    BW = Val(srl.Amount.GetValue.ToString)
                                Case "Computer Income:Color Printing"
                                    Colored = Val(srl.Amount.GetValue.ToString)
                                Case "Computer Income:Scan"
                                    Scan = Val(srl.Amount.GetValue.ToString)
                            End Select
                            If Cusage <> 0 And BW <> 0 And Colored <> 0 And Scan <> 0 Then Exit For
                        Next

                        formreport.lsvSales.Items(looper).SubItems(4).Text = Format(Cusage + BW + Colored + Scan, "#,##0.00")
                        Exit For
                    End If
                Next
            Next






        Catch ex As Exception
            MessageBox.Show(Err.Description, "COM Error")
            Return
        End Try






        'fill up report
        Dim dot As Char = Chr(149)
        formprog.Label1.Text = "creating report"
        With formreport
            .r.SelectionFont = New Font("arial", 12, FontStyle.Bold)
            .r.SelectionAlignment = HorizontalAlignment.Center
            .r.SelectedText = "Dona Report" & vbNewLine

            .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
            .r.SelectionAlignment = HorizontalAlignment.Left



            For i = 0 To .lsvSales.Items.Count - 1
                Application.DoEvents()

                formprog.ProgressBar1.Value = Int((i / .lsvSales.Items.Count) * 100 / 4) + 75
                formprog.Text = "progress - " & Int((i / .lsvSales.Items.Count) * 100 / 4) + 75 & "%"
                formprog.Refresh()
                formprog.ProgressBar1.Refresh()
                If Val(Replace(.lsvSales.Items(i).SubItems(1).Text, ",", "")) > nudLess.Value Then

                    .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
                    .r.SelectionAlignment = HorizontalAlignment.Left

                    .r.SelectionColor = Color.Black
                    .r.SelectedText += dot & " "

                    .r.SelectionColor = Color.OrangeRed
                    .r.SelectedText += .lsvSales.Items(i).SubItems(0).Text & " less value(" & .lsvSales.Items(i).SubItems(1).Text & ") has exceeded the limit(" & nudLess.Value & ")." & vbNewLine
                End If
                If Val(Replace(.lsvSales.Items(i).SubItems(2).Text, ",", "")) > nudDiscount.Value Then
                    .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
                    .r.SelectionAlignment = HorizontalAlignment.Left

                    .r.SelectionColor = Color.Black
                    .r.SelectedText += dot & " "

                    .r.SelectionColor = Color.OrangeRed
                    .r.SelectedText += .lsvSales.Items(i).SubItems(0).Text & " discount value(" & .lsvSales.Items(i).SubItems(2).Text & ") has exceeded the limit(" & nudDiscount.Value & ")." & vbNewLine
                End If
                If Val(Replace(.lsvSales.Items(i).SubItems(3).Text, ",", "")) < nudLowSale.Value Then

                    If (.lsvSales.Items(i).SubItems(3).Text <> "-----") And (.lsvSales.Items(i).SubItems(4).Text <> "-----") Then
                        .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
                        .r.SelectionAlignment = HorizontalAlignment.Left

                        .r.SelectionColor = Color.Black
                        .r.SelectedText += dot & " "

                        .r.SelectionColor = Color.Black
                        .r.SelectedText += .lsvSales.Items(i).SubItems(0).Text & " low sales(" & .lsvSales.Items(i).SubItems(3).Text & ")." & vbNewLine
                    End If
                End If
                If Val(Replace(.lsvSales.Items(i).SubItems(3).Text, ",", "")) <> Val(Replace(.lsvSales.Items(i).SubItems(4).Text, ",", "")) Then
                    .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
                    .r.SelectionAlignment = HorizontalAlignment.Left

                    .r.SelectionColor = Color.Black
                    .r.SelectedText += dot & " "

                    .r.SelectionColor = Color.Red
                    .r.SelectedText += .lsvSales.Items(i).SubItems(0).Text & " sales are not the saformmain(" & .lsvSales.Items(i).SubItems(3).Text & "<>" & .lsvSales.Items(i).SubItems(4).Text & ")." & vbNewLine
                End If
                If Val(Replace(.lsvSales.Items(i).SubItems(3).Text, ",", "")) = 0 Or Val(Replace(.lsvSales.Items(i).SubItems(4).Text, ",", "")) = 0 Then
                    .r.SelectionFont = New Font("arial", 10, FontStyle.Regular)
                    .r.SelectionAlignment = HorizontalAlignment.Left

                    .r.SelectionColor = Color.Black
                    .r.SelectedText += dot & " "

                    .r.SelectionColor = Color.Red
                    .r.SelectedText += .lsvSales.Items(i).SubItems(0).Text & " no sales for this day." & vbNewLine
                End If
            Next
            Application.DoEvents()
            formprog.ProgressBar1.Value = 100
            formprog.Text = "progress - " & "100" & "%"
            formprog.Refresh()
            formprog.ProgressBar1.Refresh()
            formprog.Hide()
            formreport.TopMost = True
            formreport.del = del
            MsgBox("Report Finished", MsgBoxStyle.SystemModal, "ROBOTICBOX")
        End With

    End Sub



    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub dtpStart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpend.ValueChanged, dtpstart.ValueChanged
        If dtpstart.Value > dtpend.Value Then dtpend.Value = dtpstart.Value
        save()
    End Sub


    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        F.SelectedPath = Application.StartupPath
        F.Description = "Locate *.mni folder"
        F.ShowDialog()
        txtmni.Text = F.SelectedPath
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        O.InitialDirectory = Application.StartupPath
        O.Title = "Locate Quickbooks File"
        O.ShowDialog()
        txtqb.Text = O.FileName
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer
        For i = 2 To 0 Step -1
            MsgBox(i)
        Next

    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        MsgBox("Leave the text box blank and it will read the opened qb file.")

    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        formreport.Show()
    End Sub
End Class
