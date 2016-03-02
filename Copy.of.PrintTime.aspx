<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 578px;
        }
    </style>
</head>

<script type="text/javascript">
function winprint() {
html= WeekTotal.innerHTML;
var win = window.open('','','width=800,height=600');
win.document.open("text/html","replace");
	win.document.write("<html><body onload='javascript:window.print()'>" + html + "</body></html>");
	win.document.close();
}
</script>
<script runat="server" >
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
</script>
<body>
    <form id="form1" runat="server">
            <table style="width:100%;">
            <tr>
                <td class="style1">
        <br />
        <asp:DropDownList ID="Name" runat="server" 
            DataSourceID="AccessDataSource1" DataTextField="LastName" 
            DataValueField="LastName">
        </asp:DropDownList>
        &nbsp;
                
&nbsp;<asp:AccessDataSource ID="AccessDataSource2" runat="server" 
            DataFile="~/time.mdb" 
            SelectCommand="SELECT [LastName], [EmployeesID], [Pass] FROM [Employees]">
        </asp:AccessDataSource>
                    </td>
                    <td>
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
                    <td class="style1">
    <asp:ListView ID="TimeView" runat="server" DataSourceID="AccessDataSource1">
        <AlternatingItemTemplate>
            <tr style="background-color: #FFFFFF;color: #284775;">
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
        </AlternatingItemTemplate>
        <LayoutTemplate>
            <table runat="server">
                <tr runat="server">
                    <td runat="server">
                        <table ID="itemPlaceholderContainer" runat="server" border="1" 
                            style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;font-family: Verdana, Arial, Helvetica, sans-serif;">
                            <tr runat="server" style="background-color: #E0FFFF;color: #333333;">
                                <th runat="server">
                                    UseDate</th>
                                <th runat="server">
                                    ClockIn</th>
                                <th runat="server">
                                    ClockOut</th>
                                <th runat="server">
                                    TotalTime</th>
                                <th runat="server">
                                    OutReason</th>
                                <th runat="server">
                                    Location</th>
                                <th runat="server">
                                    LastName</th>
                            </tr>
                            <tr ID="itemPlaceholder" runat="server">
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr runat="server">
                    <td runat="server" 
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
            <table runat="server" 
                style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;">
                <tr>
                    <td>
                        No data was returned.</td>
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
        </ItemTemplate>
    </asp:ListView>
    <asp:AccessDataSource ID="AccessDataSource1" runat="server" 
        DataFile="~/time.mdb" 
        SelectCommand="SELECT [UseDate], [ClockIn], [ClockOut], [TotalTime], [OutReason], [Location], [LastName] FROM [Time]">
    </asp:AccessDataSource>
     </td>
                </tr>
            </table>
   
    </form>
</body>
</html>
