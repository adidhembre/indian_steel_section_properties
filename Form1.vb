Imports System.Data.SQLite

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initialise()
        'Dim q As String = "SELECT * From Sections"
        'DataGridView1.DataSource = LoadDatabase(q)
    End Sub

    Public Sub Initialise()
        RadioButton1.Checked = True
        ComboBox10.Visible = False
        ComboBox11.Visible = False
        GroupBox3.Visible = False
        ComboBox3.SelectedIndex = 0
        ComboBox8.SelectedIndex = 0
        ComboBox14.SelectedIndex = 0
        ComboBox2.Enabled = False
        TextBox1.Enabled = False
        Button3.Enabled = False
        Dim chk1 As New List(Of Integer) From {4, 7}
        Dim chk2 As New List(Of Integer) From {0, 1, 2, 3, 4, 5, 6}
        For Each item As Integer In chk1
            CheckedListBox1.SetItemChecked(item, True)
        Next
        For Each item As Integer In chk2
            CheckedListBox2.SetItemChecked(item, True)
        Next
        checkListBox_Query()
    End Sub

    Public Function LoadDatabase(query As String)
        Dim cs As String = "DataSource=Resources\Section_Property_Database.db;Version=3;"
        Dim dt As DataTable = Nothing
        Dim ds As New DataSet
        'Dim reader As SQLiteDataReader

        Try
            Using conn As New SQLiteConnection(cs)
                Using cmd As New SQLiteCommand(query, conn)
                    conn.Open()
                    Using da As New SQLiteDataAdapter(cmd)
                        da.Fill(ds)
                        dt = ds.Tables(0)
                    End Using
                End Using
            End Using
            'obj.DataSource = dt
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
            'MsgBox("Error in Locating Database File")
            'Application.Exit()
        End Try
    End Function

    Public Sub checkListBox_Query()
        If Button3.Enabled() Then
            Dim index As Integer = ComboBox1.Items.IndexOf(ComboBox1.Text())
            If index > 0 Then
                CheckedListBox2.SetItemChecked(index - 1, True)
            End If
        End If
        Dim ch1 As Integer = CheckedListBox1.CheckedItems.Count
        Dim ch2 As Integer = CheckedListBox2.CheckedItems.Count
        If ch1 <> 0 And ch2 <> 0 Then
            Dim cheq As New System.Text.StringBuilder
            Dim count As Integer = 1
            cheq.Append("SELECT Name, ")
            For Each item In CheckedListBox2.CheckedItems
                Dim n As Integer = CheckedListBox2.Items.IndexOf(item)
                If n = 0 Then
                    cheq.Append("W")
                ElseIf n = 1 Then
                    cheq.Append("A")
                ElseIf n = 2 Then
                    cheq.Append("Ixx")
                ElseIf n = 3 Then
                    cheq.Append("Iyy")
                ElseIf n = 4 Then
                    cheq.Append("Zxx")
                ElseIf n = 5 Then
                    cheq.Append("Zyy")
                ElseIf n = 6 Then
                    cheq.Append("Zp")
                Else
                    MsgBox("Wrong Entry Found")
                End If
                If count < ch2 Then
                    cheq.Append(", ")
                    count = count + 1
                Else
                    cheq.Append(" FROM ")
                    Dim unit As String = ComboBox3.SelectedItem.ToString
                    If unit = "mm" Then
                        cheq.Append("SectionsInMM")
                    ElseIf unit = "cm" Then
                        cheq.Append("SectionsInCm")
                    ElseIf unit = "m" Then
                        cheq.Append("SectionsInM")
                    Else
                        MsgBox("Unit Is Not Defined")
                    End If
                    cheq.Append(" WHERE ")
                End If
            Next

            If Button3.Enabled() Then
                Dim var As String = ""
                Dim syb As String
                Dim index1 As Integer = ComboBox1.Items.IndexOf(ComboBox1.Text())
                Dim index2 As Integer = ComboBox2.Items.IndexOf(ComboBox2.Text())
                If index1 = 1 Then
                    var = "W"
                ElseIf index1 = 2 Then
                    var = "A"
                ElseIf index1 = 3 Then
                    var = "Ixx"
                ElseIf index1 = 4 Then
                    var = "Iyy"
                ElseIf index1 = 5 Then
                    var = "Zxx"
                ElseIf index1 = 6 Then
                    var = "Zyy"
                ElseIf index1 = 7 Then
                    var = "Zp"
                Else
                    'do nothing
                End If
                'None (W) (A)(Ixx)(Iyy) (Zxx)(Zyy)(Zp)
                If index1 > 0 And index2 > 0 Then
                    syb = ComboBox2.Text()
                    cheq.Append(var & " " & syb & " " & CDbl(TextBox1.Text) & " AND ( ")
                Else
                    'do nothing
                End If
                'None > < = >= <=
            Else
                'do nothing
            End If

            count = 1
            For Each item In CheckedListBox1.CheckedItems
                Dim n As Integer = CheckedListBox1.Items.IndexOf(item)
                If n = 0 Then
                    cheq.Append("subType = 'EQ'")
                ElseIf n = 1 Then
                    cheq.Append("subType = 'UE'")
                ElseIf n = 2 Then
                    cheq.Append("subType = 'ISJC'")
                ElseIf n = 3 Then
                    cheq.Append("subType = 'ISLC'")
                ElseIf n = 4 Then
                    cheq.Append("subType = 'ISMC'")
                ElseIf n = 5 Then
                    cheq.Append("subType = 'ISJB'")
                ElseIf n = 6 Then
                    cheq.Append("subType = 'ISLB'")
                ElseIf n = 7 Then
                    cheq.Append("subType = 'ISMB'")
                ElseIf n = 8 Then
                    cheq.Append("subType = 'ISWB'")
                ElseIf n = 9 Then
                    cheq.Append("subType = 'ISHB'")
                ElseIf n = 10 Then
                    cheq.Append("subType = 'ISNT'")
                ElseIf n = 11 Then
                    cheq.Append("subType = 'ISHT'")
                ElseIf n = 12 Then
                    cheq.Append("subType = 'ISST'")
                ElseIf n = 13 Then
                    cheq.Append("subType = 'ISLT'")
                ElseIf n = 14 Then
                    cheq.Append("subType = 'ISJT'")
                ElseIf n = 15 Then
                    cheq.Append("type = 'UD'")
                Else
                    MsgBox("Wrong Entry Found")
                End If
                If count < ch1 Then
                    cheq.Append(" OR ")
                    count = count + 1
                Else
                    If Button3.Enabled() Then
                        cheq.Append(" )")
                    End If
                    cheq.Append(" ;")
                End If
            Next
            'MsgBox(cheq.ToString())
            DataGridView1.DataSource = LoadDatabase(cheq.ToString())
        Else
            'MsgBox("Both Filters Can not be Null")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim thisText As String = "<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: arial, sans - serif;
  font-size:12px;
  border-collapse: collapse;
  width: 500px;
}

td, th {
  border: 1px solid #dddddd;
  text-align: Left;
  padding: 8px;
}
#r {
text-align:Right;
}
}
#c {
text-align:center;
}
.c{
margin - Left: auto;
margin-Right: auto;
margin-Top: auto;
}
</style>
</head>
<body>

<table class='c'>
  <tr>
        <th colspan ='4' id='c'>ISA 70705</th>
  </tr>
        <tr>
        <th id ='c'>Symbol</th>
    <th id='c'>Description</th>
        <th id ='c'>Value</th>
    <th id='c'>Unit</th>
          </tr>
  <tr>
        <td id ='c'>l1</td>
    <td>Leg - 1</td>
        <td id ='r'>70</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>l2</td>
    <td>Leg - 2</td>
        <td id ='r'>" + CStr(70 + 50) + "</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>t</td>
    <td>Thickness</td>
        <td id ='r'>5</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>A</td>
    <td>Cross Sectional Area</td>
        <td id ='r'>677</td>
    <td>mm<sup>2</sup></td>
          </tr>
  <tr>
        <td id ='c'>W</td>
    <td>Weight per unit length in m</td>
        <td id ='r'>5.3</td>
    <td>Kg</td>
          </tr>
  <tr>
        <td id ='c'>C<sub>xx</sub></td>
    <td>Centre of gravity in x direction</td>
        <td id ='r'>18.9</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>C<sub>yy</sub></td>
    <td>Centre of gravity in y direction</td>
        <td id ='r'>18.9</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>e<sub>xx</sub></td>
    <td>Distance of extreme fibre in x direction</td>
        <td id ='r'>51.1</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>e<sub>yy</sub></td>
    <td>Distance of extreme fibre in y direction</td>
        <td id ='r'>51.1</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>I<sub>xx</sub></td>
    <td>Moment of Inertia in x direction</td>
        <td id ='r'>311000</td>
    <td>mm<sup>4</sup></td>
          </tr>
  <tr>
        <td id ='c'>I<sub>yy</sub></td>
    <td>Moment of Inertia in y direction</td>
        <td id ='r'>311000</td>
    <td>mm<sup>4</sup></td>
          </tr>
  <tr>
        <td id ='c'>I<sub>uu</sub></td>
    <td>Moment of Inertia in u direction</td>
        <td id ='r'>498000</td>
    <td>mm<sup>4</sup></td>
          </tr>
  <tr>
        <td id ='c'>I<sub>vv</sub></td>
    <td>Moment of Inertia in v direction</td>
        <td id ='r'>125000</td>
    <td>mm<sup>4</sup></td>
          </tr>
  <tr>
        <td id ='c'>r<sub>xx</sub></td>
    <td>Radius of gyration in x direction</td>
        <td id ='r'>21.5</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>r<sub>yy</sub></td>
    <td>Radius of gyration in y direction</td>
        <td id ='r'>21.5</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>r<sub>uu</sub></td>
    <td>Radius of gyration in u direction</td>
        <td id ='r'>27.1</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>r<sub>vv</sub></td>
    <td>Radius of gyration in v direction</td>
        <td id ='r'>13.6</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>Z<sub>xx</sub></td>
    <td>Modulus of section in x direction</td>
        <td id ='r'>6100</td>
    <td>mm<sup>3</sup></td>
          </tr>
  <tr>
        <td id ='c'>Z<sub>yy</sub></td>
    <td>Modulus of section in y direction</td>
        <td id ='r'>6100</td>
    <td>mm<sup>3</sup></td>
          </tr>
  <tr>
        <td id ='c'>r<sub>1</sub></td>
    <td>Radius at Root</td>
        <td id ='r'>7</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>r<sub>2</sub></td>
    <td>Radius at Toe</td>
        <td id ='r'>4.5</td>
    <td>mm</td>
          </tr>
  <tr>
        <td id ='c'>I<sub>xy</sub></td>
    <td>Product of Inertia</td>
        <td id ='r'>184000</td>
    <td>mm<sup>4</sup></td>
          </tr>
</table>
</body>
</html>
"
        WebBrowser1.DocumentText = thisText
        WebBrowser1.Update()
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        checkListBox_Query()
    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        checkListBox_Query()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        checkListBox_Query()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim val As String = ComboBox1.Text
        Dim index As Integer = ComboBox1.Items.IndexOf(val)
        'None (W) (A)(Ixx)(Iyy) (Zxx)(Zyy)(Zp)
        If index = 0 Then
            ComboBox2.Enabled = False
            TextBox1.Enabled = False
            Button3.Enabled = False
        Else
            ComboBox2.Enabled = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim val As String = ComboBox2.Text
        Dim index As Integer = ComboBox2.Items.IndexOf(val)
        'None > < = >= <=
        If index = 0 Then
            TextBox1.Enabled = False
            Button3.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Dim val As Double = CDbl(TextBox1.Text)
            Button3.Enabled = True
        Catch ex As Exception
            Button3.Enabled = False
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        checkListBox_Query()
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Dim val As String = ComboBox9.Text
        Dim index As Integer = ComboBox9.Items.IndexOf(val)
        If index = 0 Then
            ComboBox10.Visible = False
            ComboBox11.Visible = False
            GroupBox3.Visible = False
        ElseIf index > 0 Then
            GroupBox3.Visible = False
            ComboBox10.Visible = True
            ComboBox11.Visible = False
            ComboBox10.Items.Clear()
            If index = 1 Then
                ComboBox10.Items.Add("Back To Back Angle")
                ComboBox10.Items.Add("Back To Back Channel")
                ComboBox10.Items.Add("Toe To Toe Channel")
                ComboBox10.Items.Add("Double I section")
                ComboBox10.Items.Add("I Section with Plate")
                ComboBox10.Items.Add("C Section with Plate")
            ElseIf index = 2 Then
                ComboBox10.Items.Add("I Section")
                ComboBox10.Items.Add("C Section")
                ComboBox10.Items.Add("T Section")
                ComboBox10.Items.Add("Rectangular")
                ComboBox10.Items.Add("Circular")
            End If
        Else
            ComboBox10.Visible = False
            ComboBox11.Visible = False
        End If
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        Dim subType As String = ComboBox10.Text
        Dim query As String = "SELECT Name FROM SectionsInMM WHERE subType = '" & subType & "';"
        Dim data As DataTable = LoadDatabase(query)
        If RadioButton1.Checked() Or RadioButton2.Checked() Then
            ComboBox11.Visible = False
            Calculate.Input()
        Else
            ComboBox11.Visible = True
            ComboBox11.Items.Clear()
            For Each item In data.Rows
                ComboBox11.Items.Add(item.ToString())
            Next
        End If
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        Calculate.Input()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        radioButton_Changed()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        radioButton_Changed()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        radioButton_Changed()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        radioButton_Changed()
    End Sub

    Private Sub radioButton_Changed()
        ComboBox10.Visible = False
        ComboBox11.Visible = False
        ComboBox9.SelectedIndex = 0
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        Dim text As String = ComboBox12.SelectedItem()
        If text = "Symetric" Or text = "Solid" Then
            TextBox6.Enabled = False
        Else
            TextBox6.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Calculate.Solve()
    End Sub

End Class

Public Class Calculate

    Public Shared Sub Input()
        HideAll()
        Form1.GroupBox3.Visible = True
        Form1.Label9.Text = "Name"
        Form1.Label9.Visible = True
        Form1.TextBox2.Visible = True
        If Form1.RadioButton3.Checked Or Form1.RadioButton4.Checked Then
            Form1.TextBox2.Text = Form1.ComboBox11.Text
            Form1.TextBox2.Enabled = False
        ElseIf Form1.RadioButton2.Checked Then
            Form1.TextBox2.Text = ""
            Form1.TextBox2.Enabled = True
        Else
            Form1.TextBox2.Text = "Not Required"
            Form1.TextBox2.Enabled = False
        End If
        If Form1.ComboBox9.Items.IndexOf(Form1.ComboBox9.Text) = 1 Then
            Form1.Label10.Text = "Spacing (S)"
            Form1.Label10.Visible = True
            Form1.TextBox3.Text = ""
            Form1.TextBox3.Visible = True
            Form1.Label16.Text = "Member"
            Form1.Label16.Visible = True
            If Form1.ComboBox10.Text = "Back To Back Angle" Then
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'A' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
                Form1.ComboBox13.Items.Clear()
                Form1.ComboBox13.Items.Add("Longer")
                Form1.ComboBox13.Items.Add("Shorter")
                Form1.ComboBox13.SelectedIndex = 0
                Form1.ComboBox13.Visible = True
                Form1.Label17.Text = "Connected Leg"
                Form1.Label17.Visible = True
            ElseIf Form1.ComboBox10.Text = "Back To Back Channel" Then
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'C' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "Toe To Toe Channel" Then
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'C' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "Double I section" Then
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'I' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "I Section with Plate" Then
                Form1.Label10.Visible = False
                Form1.TextBox3.Visible = False
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'I' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
                Form1.Label12.Text = "Plate Width (B)"
                Form1.Label12.Visible = True
                Form1.TextBox5.Text = ""
                Form1.TextBox5.Visible = True
                Form1.Label13.Text = "Plate Thick. (t)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.ComboBox13.Items.Clear()
                Form1.ComboBox13.Items.Add("Top & Bottom")
                Form1.ComboBox13.Items.Add("Top Only")
                Form1.ComboBox13.Items.Add("Bottom Only")
                Form1.ComboBox13.SelectedIndex = 0
                Form1.ComboBox13.Visible = True
                Form1.Label17.Text = "Plate At"
                Form1.Label17.Visible = True
            ElseIf Form1.ComboBox10.Text = "C Section with Plate" Then
                Form1.Label10.Visible = False
                Form1.TextBox3.Visible = False
                Dim data As DataTable = Form1.LoadDatabase("SELECT Name FROM SectionsInMM WHERE type = 'C' ;")
                Form1.ComboBox12.Items.Clear()
                For Each item In (From items In data.AsEnumerable() Select items.Field(Of String)(0)).ToList()
                    Form1.ComboBox12.Items.Add(item)
                Next
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
                Form1.Label12.Text = "Plate Width (B)"
                Form1.Label12.Visible = True
                Form1.TextBox5.Text = ""
                Form1.TextBox5.Visible = True
                Form1.Label13.Text = "Plate Thick. (t)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.ComboBox13.Items.Clear()
                Form1.ComboBox13.Items.Add("Top & Bottom")
                Form1.ComboBox13.Items.Add("Top Only")
                Form1.ComboBox13.Items.Add("Bottom Only")
                Form1.ComboBox13.SelectedIndex = 0
                Form1.ComboBox13.Visible = True
                Form1.Label17.Text = "Plate At"
                Form1.Label17.Visible = True
            Else
                HideAll()
            End If
        ElseIf Form1.ComboBox9.Items.IndexOf(Form1.ComboBox9.Text) = 2 Then
            If Form1.ComboBox10.Text = "I Section" Then
                Form1.Label10.Text = "Top flange width (Tbf)"
                Form1.Label10.Visible = True
                Form1.TextBox3.Text = ""
                Form1.TextBox3.Visible = True
                Form1.Label11.Text = "Top flange thick. (Ttf)"
                Form1.Label11.Visible = True
                Form1.TextBox4.Text = ""
                Form1.TextBox4.Visible = True
                Form1.Label12.Text = "Bottom flange width (Bbf)"
                Form1.Label12.Visible = True
                Form1.TextBox5.Text = ""
                Form1.TextBox5.Visible = True
                Form1.Label13.Text = "Bottom flange thick. (Btf)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.Label14.Text = "Web height (h)"
                Form1.Label14.Visible = True
                Form1.TextBox7.Text = ""
                Form1.TextBox7.Visible = True
                Form1.Label15.Text = "Web thick. (tw)"
                Form1.Label15.Visible = True
                Form1.TextBox8.Text = ""
                Form1.TextBox8.Visible = True
                Form1.Label16.Text = "Section is"
                Form1.Label16.Visible = True
                Form1.ComboBox12.Items.Clear()
                Form1.ComboBox12.Items.Add("Symetric")
                Form1.ComboBox12.Items.Add("Unsymetric")
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "C Section" Then
                Form1.Label10.Text = "Top flange width (Tbf)"
                Form1.Label10.Visible = True
                Form1.TextBox3.Text = ""
                Form1.TextBox3.Visible = True
                Form1.Label11.Text = "Top flange thick. (Ttf)"
                Form1.Label11.Visible = True
                Form1.TextBox4.Text = ""
                Form1.TextBox4.Visible = True
                Form1.Label12.Text = "Bottom flange width (Bbf)"
                Form1.Label12.Visible = True
                Form1.TextBox5.Text = ""
                Form1.TextBox5.Visible = True
                Form1.Label13.Text = "Bottom flange thick. (Btf)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.Label14.Text = "Web height (h)"
                Form1.Label14.Visible = True
                Form1.TextBox7.Text = ""
                Form1.TextBox7.Visible = True
                Form1.Label15.Text = "Web thick. (tw)"
                Form1.Label15.Visible = True
                Form1.TextBox8.Text = ""
                Form1.TextBox8.Visible = True
                Form1.Label16.Text = "Section is"
                Form1.Label16.Visible = True
                Form1.ComboBox12.Items.Clear()
                Form1.ComboBox12.Items.Add("Symetric")
                Form1.ComboBox12.Items.Add("Unsymetric")
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "T Section" Then
                Form1.Label10.Text = "Flange width (bf)"
                Form1.Label10.Visible = True
                Form1.TextBox3.Text = ""
                Form1.TextBox3.Visible = True
                Form1.Label11.Text = "flange thick. (tf)"
                Form1.Label11.Visible = True
                Form1.TextBox4.Text = ""
                Form1.TextBox4.Visible = True
                Form1.Label12.Text = "Web height (h)"
                Form1.Label12.Visible = True
                Form1.TextBox5.Text = ""
                Form1.TextBox5.Visible = True
                Form1.Label13.Text = "Web thick. (tw)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
            ElseIf Form1.ComboBox10.Text = "Rectangular" Then
                Form1.Label10.Text = "Width (B)"
                Form1.Label10.Visible = True
                Form1.TextBox3.Text = ""
                Form1.TextBox3.Visible = True
                Form1.Label11.Text = "Depth (D)"
                Form1.Label11.Visible = True
                Form1.TextBox4.Text = ""
                Form1.TextBox4.Visible = True
                Form1.Label13.Text = "Thick. (t)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.Label16.Text = "Section is"
                Form1.Label16.Visible = True
                Form1.ComboBox12.Items.Clear()
                Form1.ComboBox12.Items.Add("Solid")
                Form1.ComboBox12.Items.Add("Hollow")
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            ElseIf Form1.ComboBox10.Text = "Circular" Then
                Form1.Label10.Text = "Outer Dia (D)"
                Form1.Label10.Visible = True
                Form1.TextBox3.Text = ""
                Form1.TextBox3.Visible = True
                Form1.Label13.Text = "Thick. (t)"
                Form1.Label13.Visible = True
                Form1.TextBox6.Text = ""
                Form1.TextBox6.Visible = True
                Form1.Label16.Text = "Section is"
                Form1.Label16.Visible = True
                Form1.ComboBox12.Items.Clear()
                Form1.ComboBox12.Items.Add("Solid")
                Form1.ComboBox12.Items.Add("Hollow")
                Form1.ComboBox12.SelectedIndex = 0
                Form1.ComboBox12.Visible = True
            End If
        End If
        Form1.Button1.Visible = True
    End Sub

    Public Shared Sub HideAll()
        Form1.Label9.Visible = False
        Form1.Label10.Visible = False
        Form1.Label11.Visible = False
        Form1.Label12.Visible = False
        Form1.Label13.Visible = False
        Form1.Label14.Visible = False
        Form1.Label15.Visible = False
        Form1.Label16.Visible = False
        Form1.Label17.Visible = False
        Form1.TextBox2.Visible = False
        Form1.TextBox3.Visible = False
        Form1.TextBox4.Visible = False
        Form1.TextBox5.Visible = False
        Form1.TextBox6.Visible = False
        Form1.TextBox7.Visible = False
        Form1.TextBox8.Visible = False
        Form1.ComboBox12.Visible = False
        Form1.ComboBox13.Visible = False
    End Sub

    Public Shared Function Solve()
        Dim unit As String = Form1.ComboBox14.Text
        Dim Table As String
        If unit = "mm" Then
            Table = "SectionsInMM"
        ElseIf unit = "cm" Then
            Table = "SectionsInCM"
        ElseIf unit = "m" Then
            Table = "SectionsInM"
        Else
            Table = Nothing
        End If
        If Form1.ComboBox10.Text = "Back To Back Angle" Then
            Dim query As String = "SELECT Leg1, Leg2,W,A Cxx,Cyy,Ixx,Iyy FROM" & Table & "WHERE Name = '" & Form1.ComboBox12.Text & "';"
            Dim data As DataTable = Form1.LoadDatabase(query)
            Dim result = data.AsEnumerable()
            Dim Leg1, Leg2, W, A, Cxx, Cyy, Ixx, Iyy As Decimal
            Dim spacing As Decimal = CDec(Form1.TextBox3.Text)
            For Each item In result
                Leg1 = item.Field(Of Decimal)(0)
                Leg2 = item.Field(Of Decimal)(1)
                W = item.Field(Of Decimal)(2)
                A = item.Field(Of Decimal)(3)
                Cxx = item.Field(Of Decimal)(4)
                Cyy = item.Field(Of Decimal)(5)
                Ixx = item.Field(Of Decimal)(6)
                Iyy = item.Field(Of Decimal)(7)
            Next
            Dim Cleg, Oleg, C1, C2, I1, I2, e1, e2, NI1, NI2 As Decimal
            If Form1.ComboBox13.Text = "Longer" Then
                Cleg = Leg1
                Oleg = Leg2
                C1 = Cxx
                C2 = Cyy
                I1 = Ixx
                I2 = Iyy
            Else
                Cleg = Leg2
                Oleg = Leg1
                C1 = Cyy
                C2 = Cxx
                I1 = Iyy
                I2 = Ixx
            End If
            e1 = CDec(Math.Max(C1, Cleg - C1))
            e2 = C2
            NI1 = 2 * I1
            NI2 = 2 * I2 + 2 * (A * CDec(Math.Pow(C2, 2)))
            Dim senddata As List(Of Decimal) = ({(2 * W),
                2 * A,
                C1,
                Oleg + spacing / 2,
                e1,
                e2,
                NI1,
                NI2,
                NI1 / e1,
                NI2 / e2
                }).ToList
            Return senddata

        ElseIf Form1.ComboBox10.Text = "Back To Back Channel" Then
            Dim query As String = "SELECT bf,W,A Cxx,Cyy,Ixx,Iyy FROM" & Table & "WHERE Name = '" & Form1.ComboBox12.Text & "';"
            Dim data As DataTable = Form1.LoadDatabase(query)
            Dim result = data.AsEnumerable()
            Dim Leg1, Leg2, W, A, Cxx, Cyy, Ixx, Iyy As Decimal
            Dim spacing As Decimal = CDec(Form1.TextBox3.Text)
            For Each item In result
                Leg1 = item.Field(Of Decimal)(0)
                Leg2 = item.Field(Of Decimal)(1)
                W = item.Field(Of Decimal)(2)
                A = item.Field(Of Decimal)(3)
                Cxx = item.Field(Of Decimal)(4)
                Cyy = item.Field(Of Decimal)(5)
                Ixx = item.Field(Of Decimal)(6)
                Iyy = item.Field(Of Decimal)(7)
            Next
            Dim Cleg, Oleg, C1, C2, I1, I2, e1, e2, NI1, NI2 As Decimal
            If Form1.ComboBox13.Text = "Longer" Then
                Cleg = Leg1
                Oleg = Leg2
                C1 = Cxx
                C2 = Cyy
                I1 = Ixx
                I2 = Iyy
            Else
                Cleg = Leg2
                Oleg = Leg1
                C1 = Cyy
                C2 = Cxx
                I1 = Iyy
                I2 = Ixx
            End If
            e1 = CDec(Math.Max(C1, Cleg - C1))
            e2 = C2
            NI1 = 2 * I1
            NI2 = 2 * I2 + 2 * (A * CDec(Math.Pow(C2, 2)))
            Dim senddata As List(Of Decimal) = ({(2 * W),
                2 * A,
                C1,
                Oleg + spacing / 2,
                e1,
                e2,
                NI1,
                NI2,
                NI1 / e1,
                NI2 / e2
                }).ToList
            Return senddata
        End If
    End Function

End Class
