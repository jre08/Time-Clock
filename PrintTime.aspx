<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Web.UI" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<script type="text/javascript">
function winprint() {
html = 'Time Sheet for <b>' + document.getElementById('Name').value + '</b></br>' + document.getElementById('printview').innerHTML;
var win = window.open('','','width=800,height=600');
win.document.open("text/html","replace");
	win.document.write("<html><body onload='javascript:window.print()'>" + html + "</body></html>");
	win.document.close();
}
</script>

<SCRIPT Runat="Server">

    Sub page_load(ByVal srvSrc As Object, ByVal Args As GridViewUpdateEventArgs)
        'AccessDataSource2.SelectCommand = "SELECT ID, UseDate, ClockIn, ClockOut, TotalTime, OutReason, Location, LastName FROM [Time]WHERE lastname = '" & Name.SelectedValue & "'" & " ORDER BY ID DESC, LastName"
    End Sub
    
    Public Args As String

    
    Sub Show_Record(ByVal Src As Object, ByVal Args As EventArgs)
        Dim startnow As DateTime = StartDAte.Text
        Dim endnow As DateTime = EndDate.Text
        Dim s As String
        Dim e As String
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim Password As String
        
        DBConnection = New OleDbConnection( _
"Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        SQLString = "SELECT Pass FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        Password = DBCommand.ExecuteScalar()
        
        If Not Password = Passbox.Text Then
            Message.Visible = True
            Message.Text = "Your password was wrong "
            TimeView.Visible = False
        Else
            Message.Visible = False
            TimeView.Visible = True
            s = startnow.ToString("yyyy MM dd")
            e = endnow.ToString("yyyy MMM dd")
            AccessDataSource2.SelectCommand = "SELECT * FROM [TIME] WHERE lastname = '" & Name.SelectedValue & "' And UseDate Between '" & s & "' And  '" & e & "' order by UseDate DESC"
        End If
    End Sub
    
   
   
    Protected Sub EditGrid_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim startnow As DateTime = StartDAte.Text
        Dim endnow As DateTime = EndDate.Text
        Dim s As String
        Dim n As String
        
        s = startnow.ToString("yyyy MMM dd")
        n = endnow.ToString("yyyy MMM dd")
        AccessDataSource2.SelectCommand = "SELECT * FROM [TIME] WHERE lastname = '" & Name.SelectedValue & "' And UseDate Between '" & s & "' And  '" & n & "' Order by UseDate DESC"
        TimeView.DataBind()
    End Sub
    
    Sub Show_All(ByVal sender As Object, ByVal e As System.EventArgs)
        
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim Count As String
        Dim Password As String
        
        DBConnection = New OleDbConnection( _
"Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        SQLString = "SELECT Pass FROM Employees WHERE LastName = '" & Name.SelectedValue & "' "
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        Password = DBCommand.ExecuteScalar()
        
        If Not Password = Passbox.Text Then
            Message.Visible = True
            Message.Text = "Your password was wrong "
            TimeView.Visible = False
        Else
            TimeView.Visible = True
            Message.Visible = False
            DBConnection = New OleDbConnection( _
       "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & Server.MapPath("time.mdb"))
            DBConnection.Open()
            AccessDataSource2.SelectCommand = "Select * from [time] Where LastName = '" & Name.SelectedValue & "' Order by UseDate DESC"
            TimeView.DataBind()
            SQLString = "Select Count(*) from [time] Where LastName = '" & Name.SelectedValue & "'"
            DBCommand = New OleDbCommand(SQLString, DBConnection)
            Count = DBCommand.ExecuteScalar
        End If
    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If EndDate.Text = "" And StartDAte.Text = "" Then
            StartDAte.Text = Calendar.SelectedDate
        ElseIf StartDAte.Text > "0" And EndDate.Text = "" Then
            EndDate.Text = Calendar.SelectedDate
        Else
            StartDAte.Text = Calendar.SelectedDate
            EndDate.Text = ""
        End If
    End Sub

</SCRIPT>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
  table#Head      {border-collapse:collapse}
  table#Head th   {font-size:11pt; background-color:#E0E0E0}
  table#Head td   {font-size:11pt}
  table#Edit      {width:35px; border-collapse:collapse}
  table#Insert    {border-collapse:collapse}
  table#Insert td {font-size:10pt}
        .style1
        {
            width: 503px;
        }
        .style2
        {
            width: 333px;
        }
        .style3
        {
            width: 280px;
        }
    </style>
    </head>

<body>
    <form id="form1" runat="server">
    <div>
    



        
        <br />
        <table style="width:100%;">
            <tr>
                <td class="style2">
                    Name:&nbsp;
        <br />
        <asp:DropDownList ID="Name" runat="server" 
            DataSourceID="AccessDataSource1" DataTextField="LastName" 
            DataValueField="LastName">
        </asp:DropDownList>
        &nbsp;
                
                    <br />
                    Password:<br />
                    <font size="5" color="#FFFFFF">
          <asp:TextBox ID="Passbox" runat="server" TextMode="Password"></asp:TextBox></font>&nbsp;<asp:AccessDataSource ID="AccessDataSource1" runat="server" 
            DataFile="time.mdb" 
            SelectCommand="SELECT [LastName], [EmployeesID], [Pass] FROM [Employees]">
        </asp:AccessDataSource>
       <h3>
           <asp:Label ID="Message" runat="server"></asp:Label></h3>
                    <p>
    <input type="button" onclick="winprint()" value="Print Time Sheet" /></p>
                    </td>
                    <td align="center" class="style3">
           <asp:Calendar ID="Calendar" runat="server" 
               onselectionchanged="Calendar1_SelectionChanged"></asp:Calendar>
                    </td>
                    <td>
        <h3>Start Date:&nbsp;             <asp:TextBox ID="StartDAte" runat="server"></asp:TextBox>
&nbsp;</h3>
                        <h3>End Date:&nbsp;&nbsp;
            <asp:TextBox ID="EndDate" runat="server"></asp:TextBox>
        </h3>
                        <h3>
        <asp:Button ID="Button10" runat="server" OnClick="Show_Record" Text="Show Dates selected" />
        
        </h3>
                        <h3>
        <asp:Button ID="Button11" runat="server" OnClick="Show_All" Text="Show All Dates" />
        </h3>
                    </td>
                </tr>
                <tr>
                    <td class="style1" colspan="2">


<div id="printview">
<asp:ListView ID="TimeView" runat="server" DataSourceID="AccessDataSource2">
        <AlternatingItemTemplate>
            <tr style="background-color: #FFFFFF;color: #284775;">
                <td>
                    <asp:Label ID="UseDateLabel" runat="server" Text='<%# FormatDatetime(Eval("UseDate"),2) %>' />
                </td>
                <td width="50">
                    <asp:Label ID="ClockInLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockIn"),3) %>' />
                </td>
                <td width="50">
                    <asp:Label ID="ClockOutLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockOut"),3) %>' />
                </td>
                <td>
                    <asp:Label ID="TotalTimeLabel" runat="server" Text='<%# Eval("TotalTime") %>' />
                </td>
                <td>
                    <asp:Label ID="OutReasonLabel" runat="server" Text='<%# Eval("OutReason") %>' />
                </td>
                <td>
                    <asp:Label ID="LocationLabel" runat="server" Text='<%# Eval("Location") %>' />
                </td>
                <td>
                    <asp:Label ID="LastNameLabel" runat="server" Text='<%# Eval("LastName") %>' />
                </td>
            </tr>
        </AlternatingItemTemplate>
        <LayoutTemplate>
            <table id="Table1" runat="server">
                <tr id="Tr1" runat="server">
                    <td id="Td1" runat="server">
                        <table ID="itemPlaceholderContainer" runat="server" border="1" 
                            style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;font-family: Verdana, Arial, Helvetica, sans-serif;">
                            <tr id="Tr2" runat="server" style="background-color: #E0FFFF;color: #333333;">
                                <th id="Th1" runat="server">
                                    UseDate</th>
                                <th id="Th2" runat="server">
                                    ClockIn</th>
                                <th id="Th3" runat="server">
                                    ClockOut</th>
                                <th id="Th4" runat="server">
                                    TotalTime</th>
                                <th id="Th5" runat="server">
                                    OutReason</th>
                                <th id="Th6" runat="server">
                                    Location</th>
                                <th id="Th7" runat="server">
                                    LastName</th>
                            </tr>
                            <tr ID="itemPlaceholder" runat="server">
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="Tr3" runat="server">
                    <td id="Td2" runat="server" 
                        style="text-align: center;background-color: #5D7B9D;font-family: Verdana, Arial, Helvetica, sans-serif;color: #FFFFFF">
                    </td>
                </tr>
            </table>
        </LayoutTemplate>
        <InsertItemTemplate>
            <tr style="">
                <td>
                    <asp:Button ID="InsertButton" runat="server" CommandName="Insert" 
                        Text="Insert" />
                    <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                        Text="Clear" />
                </td>
                <td>
                    <asp:TextBox ID="UseDateTextBox" runat="server" Text='<%# Bind("UseDate") %>' />
                </td>
                <td>
                    <asp:TextBox ID="ClockInTextBox" runat="server" Text='<%# Bind("ClockIn") %>' />
                </td>
                <td>
                    <asp:TextBox ID="ClockOutTextBox" runat="server" 
                        Text='<%# Bind("ClockOut") %>' />
                </td>
                <td>
                    <asp:TextBox ID="TotalTimeTextBox" runat="server" 
                        Text='<%# Bind("TotalTime") %>' />
                </td>
                <td>
                    <asp:TextBox ID="OutReasonTextBox" runat="server" 
                        Text='<%# Bind("OutReason") %>' />
                </td>
                <td>
                    <asp:TextBox ID="LocationTextBox" runat="server" 
                        Text='<%# Bind("Location") %>' />
                </td>
                <td>
                    <asp:TextBox ID="LastNameTextBox" runat="server" 
                        Text='<%# Bind("LastName") %>' />
                </td>
            </tr>
        </InsertItemTemplate>
        <SelectedItemTemplate>
            <tr style="background-color: #E2DED6;font-weight: bold;color: #333333;">
                <td>
                    <asp:Label ID="UseDateLabel" runat="server" Text='<%# Eval("UseDate") %>' />
                </td>
                <td>
                    <asp:Label ID="ClockInLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockIn"),3) %>' />
                </td>
                <td>
                    <asp:Label ID="ClockOutLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockOut"),3) %>' />
                </td>
                <td>
                    <asp:Label ID="TotalTimeLabel" runat="server" Text='<%# Eval("TotalTime") %>' />
                </td>
                <td>
                    <asp:Label ID="OutReasonLabel" runat="server" Text='<%# Eval("OutReason") %>' />
                </td>
                <td>
                    <asp:Label ID="LocationLabel" runat="server" Text='<%# Eval("Location") %>' />
                </td>
                <td>
                    <asp:Label ID="LastNameLabel" runat="server" Text='<%# Eval("LastName") %>' />
                </td>
            </tr>
        </SelectedItemTemplate>
        <EmptyDataTemplate>
            <table id="Table2" runat="server" 
                style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;">
                <tr>
                    <td>
                        Please verify dates are correct.  No data can be found.</td>
                </tr>
            </table>
        </EmptyDataTemplate>
        <EditItemTemplate>
            <tr style="background-color: #999999;">
                <td>
                    <asp:Button ID="UpdateButton" runat="server" CommandName="Update" 
                        Text="Update" />
                    <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                        Text="Cancel" />
                </td>
                <td>
                    <asp:TextBox ID="UseDateTextBox" runat="server" Text='<%# Bind("UseDate") %>' />
                </td>
                <td>
                    <asp:TextBox ID="ClockInTextBox" runat="server" Text='<%# Bind("ClockIn") %>' />
                </td>
                <td>
                    <asp:TextBox ID="ClockOutTextBox" runat="server" 
                        Text='<%# Bind("ClockOut") %>' />
                </td>
                <td>
                    <asp:TextBox ID="TotalTimeTextBox" runat="server" 
                        Text='<%# Bind("TotalTime") %>' />
                </td>
                <td>
                    <asp:TextBox ID="OutReasonTextBox" runat="server" 
                        Text='<%# Bind("OutReason") %>' />
                </td>
                <td>
                    <asp:TextBox ID="LocationTextBox" runat="server" 
                        Text='<%# Bind("Location") %>' />
                </td>
                <td>
                    <asp:TextBox ID="LastNameTextBox" runat="server" 
                        Text='<%# Bind("LastName") %>' />
                </td>
            </tr>
        </EditItemTemplate>
        <ItemTemplate>
            <tr style="background-color: #E0FFFF;color: #333333;">
                <td>
                    <asp:Label ID="UseDateLabel" runat="server" Text='<%# FormatDateTime(Eval("UseDate"),2) %>' />
                </td>
                <td>
                    <asp:Label ID="ClockInLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockIn"),3) %>' />
                </td>
                <td>
                    <asp:Label ID="ClockOutLabel" runat="server" Text='<%# FormatDateTime(Eval("ClockOut"),3) %>' />
                </td>
                <td>
                    <asp:Label ID="TotalTimeLabel" runat="server" Text='<%# Eval("TotalTime") %>' />
                </td>
                <td>
                    <asp:Label ID="OutReasonLabel" runat="server" Text='<%# Eval("OutReason") %>' />
                </td>
                <td>
                    <asp:Label ID="LocationLabel" runat="server" Text='<%# Eval("Location") %>' />
                </td>
                <td>
                    <asp:Label ID="LastNameLabel" runat="server" Text='<%# Eval("LastName") %>' />
                </td>
            </tr>
        </ItemTemplate>
    </asp:ListView>
	<asp:Label id="Total" >
    </div>
 <asp:AccessDataSource ID="AccessDataSource2" runat="server" 
        DataFile="time.mdb" 
        >
    </asp:AccessDataSource>
        

                    </td>
                </tr>
            </table>

        
        <br />
    </div>
    <p>
    &nbsp;</p>
    </form>
</body>
</html>