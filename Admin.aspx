<%@ Page Language="VB" %>
<%@ Import Namespace="System.Drawing" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Sub Menu1_MenuItemClick(ByVal sender As Object, ByVal e As MenuEventArgs)
        Dim index As Integer = Int32.Parse(e.Item.Value)
        MultiView1.ActiveViewIndex = index
    End Sub
    
    Sub Refresh()
        MsgBox("Hey")
        EditGrid.DataBind()
        EditGrid.DataBind()
        EditGrid.DataBind()

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link href="StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <asp:Menu
        id="Menu1"
        Orientation="Horizontal"
        StaticMenuItemStyle-CssClass="tab"
        StaticSelectedStyle-CssClass="selectedTab"
        CssClass="tabs"
        OnMenuItemClick="Menu1_MenuItemClick"
        Runat="server">
        <Items>
        <asp:MenuItem Text="Time Admin" Value="0" Selected="true" />
        <asp:MenuItem Text="Users" Value="1" />
        </Items>    
    </asp:Menu>
    
    <div class="tabContents">
    <asp:MultiView
        id="MultiView1"
        ActiveViewIndex="0"
        Runat="server">
        <asp:View ID="Admin" runat="server">
        Time Admin


        </asp:View>        
        <asp:View ID="Users" runat="server">
<h3>Edit Time Clock Users</h3>
    
    <asp:FormView id="AddForm" DataSourceID="NewUser" Runat="Server"
  InsertRowStyle-BackColor="#00EE00"
  Height="49px" Width="16px">
  
<InsertRowStyle BackColor="#00EE00"></InsertRowStyle>
  
  <HeaderTemplate>
  <table id="Head" border="1">
  <tr>
    <th><asp:Label ID="Label2" Text="Edit" Width="150px" Runat="Server"/></th>
    <th><asp:Label ID="Label3" Text="ID" Width="150px"  Runat="Server"/></th>
    <th><asp:Label ID="Label4" Text="LastName" Width="150px" Runat="Server"/></th>
    <th><asp:Label ID="Label5" Text="Password" Width="150px" Runat="Server"/></th>
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
    <td  width="150px">
      <asp:Button ID="Button4" Text="Insert" CommandName="Insert" Runat="Server"
        Font-Size="7pt" Width="35px"/>
      <asp:Button ID="Button5" Text="Cancel" CommandName="Cancel" Runat="Server"
        Font-Size="7pt" Width="35px"/></td>
    <td width="150"><asp:TextBox id="AddBookID" Runat="Server"
        Text="Auto Generated" ReadOnly="true"
        Font-Size="8pt"  MaxLength="5"/></td>
   <td width="150"><asp:TextBox id="AddLastName" Runat="Server"
          Text='<%# Bind("LastName") %>'
          Font-Size="8pt" /></td>
    <td width="150"><asp:TextBox id="AddAPass" Runat="Server"
          Text='<%# Bind("Pass") %>'
          Font-Size="8pt" /></td>
      </tr>
  </table>
  </InsertItemTemplate>

</asp:FormView>
<asp:GridView id="EditGrid" DataSourceID="EditUser" Runat="Server"
  AutoGenerateColumns="False"
  DataKeyNames="EmployeesID"
  ShowHeader="False"
  AllowPaging="True"
  EditRowStyle-BackColor="#FFFF00"
  PagerStyle-BackColor="#E0E0E0"
  RowStyle-VerticalAlign="Top"
  RowStyle-Font-Size="10pt"
  RowCommand="Refresh"
 RowStyle-Width="150">
  
<RowStyle VerticalAlign="Top" Font-Size="10pt"></RowStyle>
  
  <Columns>
  
      <asp:CommandField ShowDeleteButton="True" />
      <asp:CommandField ShowEditButton="true" >
  
  
          <ControlStyle Width="150px" />
      </asp:CommandField>
  
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label11" Text='<%# Eval("EmployeesID") %>' Runat="Server"
        Width="150"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Label ID="Label12" Text='<%# Eval("EmployeesID") %>' Runat="Server"
        Width="150"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label14" Text='<%# Eval("LastName") %>' Runat="Server"
        Width="150"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="EditLastName" Runat="Server"
        Text='<%# Bind("LastName") %>'
        Width="150" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label15" Text='<%# Eval("Pass") %>' Runat="Server"
        Width="150"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="EditPass" Runat="Server"
        Text='<%# Bind("Pass") %>'
        Width="150" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  
  </Columns>

<PagerStyle BackColor="#E0E0E0"></PagerStyle>

<EditRowStyle BackColor="Yellow"></EditRowStyle>

</asp:GridView>
        </asp:View>                
    </asp:MultiView>
    </div>
    </div>
    <asp:AccessDataSource ID="NewUser" runat="server" 
            DataFile="time.mdb" 
             InsertCommand="INSERT INTO Employees (LastName, Pass, ClockIn, UseDate)
                 VALUES (@LastName, @Pass, 0, '1/1/1900')"
            SelectCommand="SELECT * FROM [Employees] ORDER BY [LastName]"
            UpdateCommand="UPDATE Employees SET LastName=@LastName, Pass=@Pass WHERE EmployeesID=@EmployeesID"
            OnInserted="Refresh"
             />
        
        <asp:AccessDataSource id="EditUser" Runat="Server"
  DataFile="time.mdb"
  
  SelectCommand="SELECT * FROM Employees ORDER BY [LastName]"
  
  UpdateCommand="UPDATE Employees SET LastName=@LastName, Pass=@Pass
                 WHERE EmployeesID=@EmployeesID"
  DeleteCommand="Delete * From [Employees] WHERE EmployeesID = @EmployeesID"
/>
    </form>
</body>
</html>
