<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<%@ Import Namespace="System.Drawing" %>

<SCRIPT Runat="Server">

Sub Validate_Update_Data (Src As Object, Args As GridViewUpdateEventArgs)

        If Args.NewValues("LastName") = "" Then
            Args.Cancel = True
            EditMSG.Text = "&bull; Missing Name"
        End If
  
        If Args.NewValues("Pass") = "" Then
            Args.Cancel = True
            EditMSG.Text = "&bull; Missing author"
        End If
    End Sub

Sub Update_Record (Src As Object, Args As GridViewUpdatedEventArgs)

        EditMSG.Text = "&bull; Record " & Args.Keys("EmployeesID") & " updated"
  
End Sub

Sub Validate_Insert_Data (Src As Object, Args As FormViewInsertEventArgs)

        If Args.Values("LastName") = "" Then
            Args.Cancel = True
            EditMSG.Text = "&bull; Missing Name"
        End If
  
        If Args.Values("Pass") = "" Then
            Args.Cancel = True
            EditMSG.Text = "&bull; Missing Password"
        End If
    End Sub

Sub Insert_Record (Src As Object, Args As FormViewInsertedEventArgs)

  EditGrid.DataBind()
        EditMSG.Text = "&bull; Record " & Args.Values("EmployeesID") & " added"
  
End Sub

    Sub Confirm_Delete(ByVal Src As Object, ByVal Args As GridViewDeleteEventArgs)

        Args.Cancel = True
        ConfirmDelete.Visible = True
        Dim Row As GridViewRow = EditGrid.Rows(Args.RowIndex)
        Row.BackColor = Color.FromName("#FF3333")
        Row.ForeColor = Color.FromName("#FFFFFF")
        ViewState("RowIndex") = Args.RowIndex
        ViewState("EmployeesID") = Args.Keys("EmployeesID")
  
    End Sub
    Dim Delete As Parameter
    Sub Delete_Record(ByVal Src As Object, ByVal Args As EventArgs)
              
        ' EditUser.DeleteCommand = "DELETE FROM Employees WHERE EmployeesID = '" & ViewState("EmployeesID") & "'"
        EditUser.Delete()
        Dim Row As GridViewRow = EditGrid.Rows(ViewState("RowIndex"))
        Row.BackColor = Color.FromName("#FFFFFF")
        Row.ForeColor = Color.FromName("#000000")
        EditMSG.Text = "&bull; Record " & ViewState("BookID") & " deleted"
  
    End Sub

    Sub Cancel_Delete(ByVal Src As Object, ByVal Args As EventArgs)

        Dim Row As GridViewRow = EditGrid.Rows(ViewState("RowIndex"))
        Row.BackColor = Color.FromName("#FFFFFF")
        Row.ForeColor = Color.Fromname("#000000")
        ConfirmDelete.Visible = False
  
    End Sub

    
</SCRIPT>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">

        .style1
        {
            width: 70px;
        }
        .style2
        {
            width: 137px;
        }
    </style>
</head>

<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:AccessDataSource ID="NewUser" runat="server" 
            DataFile="time.mdb" 
             InsertCommand="INSERT INTO Employees (LastName, Pass, ClockIn, UseDate)
                 VALUES (@LastName, @Pass, 0, '1/1/1900')"
            SelectCommand="SELECT * FROM [Employees] ORDER BY [LastName]"
            UpdateCommand="UPDATE Employees SET LastName=@LastName, Pass=@Pass WHERE EmployeesID=@EmployeesID" />
        
        <asp:AccessDataSource id="EditUser" Runat="Server"
  DataFile="time.mdb"
  
  SelectCommand="SELECT * FROM Employees ORDER BY [LastName]"
  
  UpdateCommand="UPDATE Employees SET LastName=@LastName, Pass=@Pass
                 WHERE EmployeesID=@EmployeesID"
  DeleteCommand="Delete * From [Employees] WHERE EmployeesID = @EmployeesID"
/>

       
       <h3>Edit Time Clock Users</h3>

<asp:Label id="EditMSG" Height="25" ForeColor="Red" Runat="Server"
EnableViewState="False"/>

<asp:Label id="ConfirmDelete" Visible="False" Runat="Server"
EnableViewState="False" Height="25">
  <asp:Label ID="Label1" Text="Delete this record? " Runat="Server"
    ForeColor="Red" EnableViewState="False"/>
  <asp:Button ID="Button1" Text="Yes" OnClick="Delete_Record" Runat="Server"
    Font-Size="7pt" Width="30px"/>
  <asp:Button ID="Button2" Text="No" OnClick="Cancel_Delete" Runat="Server"
    Font-Size="7pt" Width="30px"/>
</asp:Label>

<style type="text/css">
  table#Head      {border-collapse:collapse}
  table#Head th   {font-size:11pt; background-color:#E0E0E0}
  table#Head td   {font-size:11pt}
  table#Edit      {width:35px; border-collapse:collapse}
  table#Insert    {border-collapse:collapse}
  table#Insert td {font-size:10pt}
</style>

<asp:FormView id="AddForm" DataSourceID="NewUser" Runat="Server"
  InsertRowStyle-BackColor="#00EE00"
  OnItemInserting="Validate_Insert_Data"
  OnItemInserted="Insert_Record" Height="49px" Width="16px">
  
<InsertRowStyle BackColor="#00EE00"></InsertRowStyle>
  
  <HeaderTemplate>
  <table id="Head" border="1">
  <tr>
    <th><asp:Label ID="Label2" Text="Edit" Width="80px" Runat="Server"/></th>
    <th><asp:Label ID="Label3" Text="ID" Width="80px" Runat="Server"/></th>
    <th><asp:Label ID="Label4" Text="LastName" Width="80px" Runat="Server"/></th>
    <th><asp:Label ID="Label5" Text="Password" Width="80px" Runat="Server"/></th>
    </tr>
  </table>
  </HeaderTemplate>
  
  <ItemTemplate>
  <table id="Edit" border="1">
  <tr>
    <td><asp:Button ID="Button3" Text="New" CommandName="New" Runat="Server"
        Font-Size="7pt" Width="35px"/></td>
  </tr>
  </table>
  </ItemTemplate>
  
  <InsertItemTemplate>
  <table id="Insert" border="1">
  <tr>
    <td  nowrap>
      <asp:Button ID="Button4" Text="Insert" CommandName="Insert" Runat="Server"
        Font-Size="7pt" Width="35px"/>
      <asp:Button ID="Button5" Text="Cancel" CommandName="Cancel" Runat="Server"
        Font-Size="7pt" Width="35px"/></td>
    <td><asp:TextBox id="AddBookID" Runat="Server"
        Text="Auto Generated" ReadOnly="true"
        Font-Size="8pt" Width="80" MaxLength="5"/></td>
   <td><asp:TextBox id="AddLastName" Runat="Server"
          Text='<%# Bind("LastName") %>'
          Font-Size="8pt" Width="80"/></td>
    <td><asp:TextBox id="AddAPass" Runat="Server"
          Text='<%# Bind("Pass") %>'
          Font-Size="8pt" Width="80"/></td>
      </tr>
  </table>
  </InsertItemTemplate>

</asp:FormView>
<asp:GridView id="EditGrid" DataSourceID="EditUser" Runat="Server"
  AutoGenerateColumns="False"
  DataKeyNames="EmployeesID"
  ShowHeader="False"
  AllowPaging="True"
  PageSize="10"
  EditRowStyle-BackColor="#FFFF00"
  PagerStyle-BackColor="#E0E0E0"
  RowStyle-VerticalAlign="Top"
  RowStyle-Font-Size="10pt"
  OnRowUpdating="Validate_Update_Data"
  OnRowUpdated="Update_Record"
 >
  
<RowStyle VerticalAlign="Top" Font-Size="10pt"></RowStyle>
  
  <Columns>
  
      <asp:CommandField ShowDeleteButton="True" />
  
  <asp:TemplateField 
    ItemStyle-Wrap="False">
    <ItemTemplate>
      <asp:Button ID="Button6" Text="Edit" CommandName="Edit" Runat="Server"
        Font-Size="7pt" Width="35px"/>
      
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Button ID="Button8" Text="Update" CommandName="Update" Runat="Server"
        Font-Size="7pt" Width="35px"/>
      <asp:Button ID="Button9" Text="Cancel" CommandName="Cancel" Runat="Server"
        Font-Size="7pt" Width="35px"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label11" Text='<%# Eval("EmployeesID") %>' Runat="Server"
        Width="80"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Label ID="Label12" Text='<%# Eval("EmployeesID") %>' Runat="Server"
        Width="80"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label14" Text='<%# Eval("LastName") %>' Runat="Server"
        Width="80"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="EditLastName" Runat="Server"
        Text='<%# Bind("LastName") %>'
        Width="80" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label15" Text='<%# Eval("Pass") %>' Runat="Server"
        Width="80"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="EditPass" Runat="Server"
        Text='<%# Bind("Pass") %>'
        Width="80" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  
  </Columns>

<PagerStyle BackColor="#E0E0E0"></PagerStyle>

<EditRowStyle BackColor="Yellow"></EditRowStyle>

</asp:GridView>
    </div>
    </form>
</body>
</html>
