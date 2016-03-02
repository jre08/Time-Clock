<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<script runat="server">

    
    Sub Page_load(ByVal Src As Object, ByVal Args As EventArgs)
                  
        
        
    End Sub
    
    Sub ClockIn(ByVal Src As Object, ByVal Args As EventArgs)
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim DBReader As OleDbDataReader
        Dim SQLString As String
        Dim Password As String
        Dim UseDAte As DateTime
        Dim ClockedIn As String
        Dim DateUsed As String
        Dim IDnum As String
        Message.Visible = True
        Image.Visible = False
      
	
        DBConnection = New OleDbConnection( _
    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        SQLString = "SELECT Pass FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        Password = DBCommand.ExecuteScalar()
        
        SQLString = "SELECT Clockin,UseDate FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBReader = DBCommand.ExecuteReader()
        DBReader.Read()
        
        ClockedIn = DBReader("ClockIn")
        DateUsed = DBReader("UseDate")
        
        DBReader.Close()
        If Not Password = Passbox.Text Then
            Message.Text = "Your password was wrong " & FormatDateTime(Now, DateFormat.ShortTime)
        ElseIf LocationList.SelectedValue = "" Then
            Message.Text = "Please Select a location."
        ElseIf ClockedIn > "0" And DateUsed = FormatDateTime(Now, DateFormat.ShortDate) Then
            Message.Text = "You are already clocked in for today, please clock out."
        ElseIf FormatDateTime(Now, DateFormat.ShortTime) < "07:45" Then
            Message.Text = "You cannot clock in before 7:45, please wait until then to do so."
        Else
           
            UseDAte = FormatDateTime(Today, 1)
            Dim UDate As String
            UDate = UseDAte.ToString("yyyy MM dd")
            SQLString = "INSERT INTO [time]" & _
             "(UseDate, ClockIn, ClockOut, Location, LastName)" & _
             "VALUES (" & "'" & UDate & "','" & FormatDateTime(Now(), 4) & "','" & FormatDateTime("00:00", 4) & "','" & LocationList.SelectedValue & "','" & Name.SelectedValue & "')"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBCommand.ExecuteNonQuery()
            ' SQLString = "UPDATE Employees SET " & _
            '"ClockIn='" & FormatDateTime(Now(), 4) & "'" & _
            '" WHERE LastName='" & Name.SelectedValue & "'"
            'DBCommand = New OleDbCommand(SQLString, DBConnection)
            'DBCommand.ExecuteNonQuery()
            
            'DBConnection.Close()
            
            
            SQLString = "SELECT ID FROM [time] WHERE LastName = '" & Name.SelectedValue & "' AND UseDate ='" & UDate & "' AND ClockOut = '" & FormatDateTime("00:00", 4) & "'"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBReader = DBCommand.ExecuteReader()
            DBReader.Read()
        
            IDnum = DBReader("ID")
        
            DBReader.Close()
            
            SQLString = "UPDATE Employees SET " & _
            "ClockIn='" & IDnum & "', UseDate ='" & FormatDateTime(Now, DateFormat.ShortDate) & "'" & _
             " WHERE LastName='" & Name.SelectedValue & "'"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBCommand.ExecuteNonQuery()
            
            DBConnection.Close()

            Message.Text = Name.SelectedValue + ", you have clocked in.  Thank you have a nice day!"
            Response.AppendHeader("Refresh", "5")
        End If
        
    End Sub
    
    Sub ClockOut(ByVal Src As Object, ByVal Args As EventArgs)
        
        
        ' Add these lines to the clock out 
        'elseif datepart("h", now) < "17" and document.getElementById("OutReason").value = "" then 
        'document.getElementById("message").innerHTML = Name1.value & ", Why are you clocking out?"
        
        'DON"T FORGET TO ADD REASON TO UPDATE STATEMENT
        ' TimeRecordset("OutReason") = document.getElementByID("OutReason").value
        
        
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim DBReader As OleDbDataReader
        Dim ClockIn As String
        Dim ClockOut As String
        Dim UseDate As DateTime
        Dim HourDif As String
        Dim MinDif As String
        Dim TotalTime As String
        Dim Password As String
        Dim ClockedIn As String
        Dim IDnum As String
        
        Message.Visible = True
        Image.Visible = False
        
        
        
        
        DBConnection = New OleDbConnection( _
   "Provider=Microsoft.Jet.OLEDB.4.0;" & _
   "Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        SQLString = "SELECT Pass FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        Password = DBCommand.ExecuteScalar()
        
        SQLString = "SELECT * FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBReader = DBCommand.ExecuteReader()
        DBReader.Read()
        
        IDnum = DBReader("EmployeesID")
        ClockedIn = DBReader("ClockIn")
        
        DBReader.Close()
        
        If Not Password = Passbox.Text Then
            Message.Text = "Your password was wrong"
        ElseIf ClockedIn = "0" Then
            Message.Text = "You are not clocked in, please clockin."
        Else
          
            SQLString = "Select * from [Time] " & _
                  "Where ID = " & ClockedIn & ""
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBReader = DBCommand.ExecuteReader()
            DBReader.Read()
        
            UseDate = DBReader("UseDate")
            ClockIn = DBReader("ClockIn")
            ClockOut = DBReader("ClockOut")
        
            DBReader.Close()
        
            SQLString = "UPDATE [Time] SET " & _
            "ClockOut = '" & FormatDateTime(Now(), 4) & "' " & _
            "WHERE ID = " & ClockedIn & ""
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBCommand.ExecuteNonQuery()
        
            Dim Udate As String
            Udate = UseDate.ToString("yyyy MM dd")
            SQLString = "Select * from [Time] Where UseDate='" & Udate & "'" & " and LastName=" & "'" & Name.SelectedValue & "'" & " and ClockIn=" & "'" & ClockIn & "'"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBReader = DBCommand.ExecuteReader()
            DBReader.Read()
        
            ClockIn = DBReader("ClockIN")
            ClockOut = DBReader("ClockOut")

            DBReader.Close()
        
            If DatePart("h", ClockOut) = DatePart("h", ClockIn) Then
                HourDif = 0
            ElseIf DatePart("h", ClockOut) = 0 Then
                HourDif = 24 - DatePart("h", ClockIn)
            ElseIf DatePart("h", ClockOut) < DatePart("h", ClockIn) Then
                HourDif = DatePart("h", ClockOut) - DatePart("h", ClockIn) - 1
            Else
                HourDif = DatePart("h", ClockOut) - DatePart("h", ClockIn)

            End If
        
            If DatePart("n", ClockIn) > DatePart("n", ClockOut) And HourDif = 1 Then
                HourDif = 0
                MinDif = 60 - DatePart("n", ClockIn) + DatePart("n", ClockOut)
            ElseIf DatePart("h", ClockOut) > DatePart("h", ClockIn) And DatePart("n", ClockOut) > DatePart("n", ClockIn) Then
                MinDif = DatePart("n", ClockOut) - DatePart("n", ClockIn)
                HourDif = DatePart("h", ClockOut) - DatePart("h", ClockIn)
                '2
            ElseIf DatePart("h", ClockOut) > DatePart("h", ClockIn) And DatePart("n", ClockIn) > DatePart("n", ClockOut) Then
                MinDif = 60 - DatePart("n", ClockIn) + DatePart("n", ClockOut)
                HourDif = DatePart("h", ClockOut) - DatePart("h", ClockIn) - 1
            ElseIf DatePart("n", ClockIn) = DatePart("n", ClockOut) Then
                MinDif = 0
            ElseIf DatePart("n", ClockOut) = 0 Then
                MinDif = 60 - DatePart("n", ClockIn)
            ElseIf DatePart("n", ClockIn) > DatePart("n", ClockOut) Then
                HourDif = HourDif - 1
                MinDif = DatePart("n", ClockIn) - DatePart("n", ClockOut)
            Else
                MinDif = DatePart("n", ClockOut) - DatePart("n", ClockIn)
            End If

            If Len(MinDif) = 1 Then
                MinDif = 0 & MinDif
            End If

            TotalTime = HourDif & ":" & MinDif
            SQLString = "UPDATE [Time] SET " & _
                    "TotalTime = '" & TotalTime & "' " & _
                    "WHERE ID = " & ClockedIn & ""
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBCommand.ExecuteNonQuery()
            SQLString = "UPDATE [Employees] SET ClockIN = 0, UseDate = '1/1/1900' WHERE LastName = '" & Name.SelectedValue & "'"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            DBCommand.ExecuteNonQuery()
            DBConnection.Close()
            Message.Text = Name.SelectedValue + ", you have clocked Out.  Thank you have a nice day!"
        End If
        Response.AppendHeader("Refresh", "5")
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    </head>
<body>
    <form id="form1" runat="server">
    <div align="center">
    <table id="color" bgcolor="#0000FF" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="680px" height="300px">
    <tr>
      <td  width="29%" height="300px" align="center">
          <font size="5" color="#FFFFFF">Name<br />
          </font>
          
          <asp:DropDownList ID="Name" runat="server" 
              DataSourceID="AccessDataSource1" DataTextField="LastName" 
              DataValueField="LastName">
          </asp:DropDownList>
          <asp:AccessDataSource ID="AccessDataSource1" runat="server" 
              DataFile="time.mdb" SelectCommand="SELECT [LastName] FROM [Employees]">
          </asp:AccessDataSource>
      
      
      <p><font size="5" color="#FFFFFF">Password
          <asp:TextBox ID="Passbox" runat="server" TextMode="Password"></asp:TextBox></font><br />
      
                            </p>
          <font size="5" color="#FFFFFF">Location</font><asp:RadioButtonList ID="LocationList" runat="server" Font-Size="Large" 
              ForeColor="White">
              <asp:ListItem>Marianna</asp:ListItem>
              <asp:ListItem>Crestview</asp:ListItem>
              <asp:ListItem>Panama City</asp:ListItem>
          </asp:RadioButtonList>
        <p><font size="5" color="#FFFFFF">Clock Out Reason</font><br />
        <font size="5" color="#FFFFFF">
          <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox></font>
        </p>
        <p>
            <asp:Button ID="ClkIn" runat="server" Text="Clock In" onclick="ClockIn" />
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <asp:Button ID="ClkOut" runat="server" Text="Clock Out" 
                onclick="ClockOut" /></p>
      </td>
      <td width="71%" height="171" valign="middle" align="center">
      <p><span style="height:135px; width:470px; background-color:#ffffff;">
          <asp:Label ID="Message" runat="server" Height="135px" Width="470px" 
              EnableViewState="False" Visible="False" /> 
          <asp:Image ID="Image" runat="server" ImageUrl="Logo.jpg" />
      </span>&nbsp;</p>
      <embed 
  src="clock.swf" 
  width="150"
  height="150"
  allowscriptaccess="always"
/>
      &nbsp;
      </td>
    </tr>
    </table>
</div>
    </form>
    </body>
</html>
