<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Drawing" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">




<SCRIPT Runat="Server">

    Sub page_load(ByVal srvSrc As Object, ByVal Args As GridViewUpdateEventArgs)
        NewTime.SelectCommand = "SELECT ID, UseDate, ClockIn, ClockOut, TotalTime, OutReason, Location, LastName FROM [Time]WHERE lastname = '" & Name.SelectedValue & "'" & " ORDER BY ID DESC, LastName"
    End Sub
    
    Sub Validate_Update_Data(ByVal srvSrc As Object, ByVal Args As GridViewUpdateEventArgs)

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
        EditMSG.Text = "&bull; Record " & Args.OldValues("LastName") & " updated"
        
        '*** Start TotalTime Update
        
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim DBReader As OleDbDataReader
        Dim ClockIn As String
        Dim ClockOut As String
        Dim HourDif As String
        Dim MinDif As String
        Dim TotalTime As String
        Dim IDNum As String
        
        
        DBConnection = New OleDbConnection( _
    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        
        SQLString = "UPDATE [Time] SET " & _
                "ClockIn = '" & Args.NewValues("ClockIN") & "', ClockOut='" & Args.NewValues("ClockOUt") & "'" & _
                "WHERE ID = " & Args.Keys("ID") & ""
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBCommand.ExecuteNonQuery()

        '**** NEEDS EDTING BELOW *********
        SQLString = "Select * from [time] Where ID = " & Args.Keys("ID") & ""
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBReader = DBCommand.ExecuteReader()
        DBReader.Read()
        
        
        ClockIn = DBReader("ClockIn")
        ClockOut = DBReader("ClockOut")
        IDNum = DBReader("ID")

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
                "WHERE ID = " & IDNum & ""
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBCommand.ExecuteNonQuery()
        DBConnection.Close()
        DataBind()
        
        
        
        
        
        '*** End TotalTime Update
  
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
    Public Args As String
Sub Insert_Record (Src As Object, Args As FormViewInsertedEventArgs)
        EditMSG.Text = "&bull; Record " & Args.Values("LastName") & " added"
 
        
        '*** Start TotalTime Update
        
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim DBReader As OleDbDataReader
        Dim ClockIn As String
        Dim ClockOut As String
        Dim HourDif As String
        Dim MinDif As String
        Dim TotalTime As String
        Dim IDNum As String
        
        
        DBConnection = New OleDbConnection( _
    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()

        '**** NEEDS EDTING BELOW *********
        SQLString = "Select * from [time] Where UseDate ='" & Args.Values("UseDate") & "' AND ClockIn = '" & Args.Values("ClockIN") & "' And ClockOut = '" & Args.Values("ClockOUt") & "' And LastName = '" & Args.Values("LastName") & "'"
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBReader = DBCommand.ExecuteReader()
        DBReader.Read()
        MsgBox(Args.Values("UseDate") & Chr(10) & Args.Values("ClockIN") & Chr(10) & Args.Values("ClockOUt") & Chr(10) & Args.Values("LastName"))
        
        ClockIn = DBReader("ClockIn")
        ClockOut = DBReader("ClockOut")
        IDNum = DBReader("ID")

        DBReader.Close()
        
        MsgBox(IDNum)
        
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
                "WHERE ID = " & IDNum & ""
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        DBCommand.ExecuteNonQuery()
        DBConnection.Close()
        DataBind()
        
        
        
        
        
        '*** End TotalTime Update
        
        
  
    End Sub
    
    Sub Delete_Record(ByVal Src As Object, ByVal Args As EventArgs)
        MsgBox(ViewState("RowIndex"))
        'EditTime.DeleteCommand = "DELETE FROM [Time] WHERE ID = '" & ViewState("ID") & "'"
        'EditTime.Delete()
        'Dim Row As GridViewRow = EditGrid.Rows(ViewState("RowIndex"))
        'Row.BackColor = Color.FromName("#FFFFFF")
        'Row.ForeColor = Color.FromName("#000000")
  
    End Sub

    Sub Cancel_Delete(ByVal Src As Object, ByVal Args As EventArgs)

        Dim Row As GridViewRow = EditGrid.Rows(ViewState("RowIndex"))
        Row.BackColor = Color.FromName("#FFFFFF")
        Row.ForeColor = Color.FromName("#000000")
        ConfirmDelete.Visible = False
  
    End Sub
    
    Sub Show_Record(ByVal Src As Object, ByVal Args As EventArgs)
        EditGrid.PageSize = "10"
        EditTime.SelectCommand = "SELECT * FROM [TIME] WHERE lastname = '" & Name.SelectedValue & "'"
    End Sub
    
   
   
    Protected Sub EditGrid_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        EditTime.SelectCommand = "SELECT * FROM [TIME] WHERE lastname = '" & Name.SelectedValue & "' Order by UseDate DESC"
        EditGrid.DataBind()
    End Sub
    
    Sub Show_All(ByVal sender As Object, ByVal e As System.EventArgs)
        
        Dim DBConnection As OleDbConnection
        Dim DBCommand As OleDbCommand
        Dim SQLString As String
        Dim Count As String
                
        DBConnection = New OleDbConnection( _
   "Provider=Microsoft.Jet.OLEDB.4.0;" & _
   "Data Source=" & Server.MapPath("time.mdb"))
        DBConnection.Open()
        SQLString = "Select * from [time] Where LastName = '" & Name.SelectedValue & "'"
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        SQLString = "Select Count(*) from [time] Where LastName = '" & Name.SelectedValue & "'"
        DBCommand = New OleDbCommand(SQLString, DBConnection)
        Count = DBCommand.ExecuteScalar
        EditGrid.PageSize = Count
        Label.Text = Count
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
</style>
    </head>

<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:AccessDataSource id="EditTime" Runat="Server"
  DataFile="~/time.mdb"   
  DeleteCommand="Delete FROM [TIME] WHERE ID = [@ID]"
  UpdateCommand="Update [time] Set ClockIn=@ClockIN Where ID=@ID"
  SelectCommand = "SELECT ID, UseDate, ClockIn, ClockOut, TotalTime, OutReason, Location, LastName FROM [Time] where ([LastName] = ?) ORDER BY LastName, ID ASC" CancelSelectOnNullParameter="False">
  <SelectParameters>
                       <asp:ControlParameter ControlID="Name" Name="LastName" PropertyName="SelectedValue" Type="String" />
              </SelectParameters>
  </asp:AccessDataSource>




        
        <br />
        <br />
        <asp:DropDownList ID="Name" runat="server" 
            DataSourceID="AccessDataSource1" DataTextField="LastName" 
            DataValueField="LastName">
        </asp:DropDownList>
        &nbsp;
        <asp:Button ID="Button10" runat="server" OnClick="Show_Record" Text="Button" />
        
&nbsp;<asp:AccessDataSource ID="AccessDataSource1" runat="server" 
            DataFile="~/time.mdb" 
            SelectCommand="SELECT [LastName], [EmployeesID], [Pass] FROM [Employees]">
        </asp:AccessDataSource>
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



            <br />
            <br />
    
        <asp:AccessDataSource ID="NewTime" runat="server" 
            DataFile="~/time.mdb" 
             InsertCommand="INSERT INTO [Time] (UseDate,ClockIn,ClockOut,OutReason,Location,LastName)
                 VALUES (@UseDate,@ClockIn,@ClockOut,@TotalTime,@OutReason,@Location,@LastName)"
            
             />
        


<asp:FormView id="AddForm" DataSourceID="NewTime" Runat="Server"
  InsertRowStyle-BackColor="#00EE00"
  OnItemInserted="Insert_Record" Height="49px">
  
  
  <HeaderTemplate>
  <table id="Head" border="1">
  <tr>
    <th><asp:Label ID="EditTH" Text="Edit" Width="85px" Runat="Server"/></th>
    <th><asp:Label ID="DateTH" Text="Date" Width="85px" Runat="Server"/></th>
    <th><asp:Label ID="ClockInTH" Text="Clock In" Width="85px" Runat="Server"/></th>
    <th><asp:Label ID="ClockOutTH" Text="Clock Out" Width="85px" Runat="Server"/></th>
    <th><asp:Label ID="TTimeTH" Text="Total Time" Width="85px" Runat="server" /></th>
    <th><asp:Label ID="ReasonTH" Text="Clock Out Reason" Width="85px" Runat="Server"/></th>
    <th><asp:Label ID="LocationTH" Text="Location" Width="85px" runat="server" /></th>
    <th><asp:Label ID="NameTH" Text="Name" Width="85px" Runat="Server"/></th>
    </tr>
  </table>
  </HeaderTemplate>
  
  <ItemTemplate>
  <table id="Edit" border="1">
  <tr>
    <td><asp:Button ID="Button3" Text="New" CommandName="New" Runat="Server"
        Font-Size="7pt" Width="40px"/></td>
  </tr>
  </table>
  </ItemTemplate>
  
 

</asp:FormView>


<asp:GridView id="EditGrid" DataSourceID="EditTime" Runat="Server"
  AutoGenerateColumns="False"
  DataKeyNames="ID"
  ShowHeader="False"
  AllowPaging="True"
  PageSize="10"
  EditRowStyle-BackColor="#FFFF00"
  PagerStyle-BackColor="#E0E0E0"
  RowStyle-VerticalAlign="Top"
  RowStyle-Font-Size="10pt"
  OnRowUpdated="Update_Record" onselectedindexchanged="EditGrid_SelectedIndexChanged1"
 >
  
<RowStyle VerticalAlign="Top" Font-Size="10pt"></RowStyle>
  
  <Columns> 
  <asp:TemplateField 
    ItemStyle-Wrap="False">
    <ItemTemplate>
      <asp:Button ID="Button6" Text="Edit" CommandName="Edit" Runat="Server"
        Font-Size="7pt" Width="42px"/>
      <asp:Button ID="Delete" Text="Delete" CommandName="Delete" runat="server" 
        Font-Size="7pt" Width="42px"/>
      
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Button ID="Button8" Text="Update" CommandName="Update" Runat="Server"
        Font-Size="7pt" Width="35px"/>
      <asp:Button ID="Button9" Text="Cancel" CommandName="Cancel" Runat="Server"
        Font-Size="7pt" Width="35px"/>
    </EditItemTemplate>

<ItemStyle Wrap="False"></ItemStyle>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="IDbox" Text='<%# Eval("ID") %>' Runat="Server"
        Width="85" Visible="False" />
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Label ID="IDEdit" Text='<%# Eval("ID") %>' Runat="Server"
        Width="85" Visible="False"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="UseDatebox" Text='<%# Eval("UseDate") %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:Label ID="UseDateEdit" Text='<%# Eval("UseDate") %>' Runat="Server"
        Width="85"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="ClockInbox" Text='<%# FormatDateTime(Eval("ClockIn"),3) %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="ClockInEdit" Runat="Server"
        Text='<%# Bind("ClockIN") %>'
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="ClockOutbox" Text='<%#FormatDateTime(Eval("ClockOut"),3) %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="ClockOutEdit" Runat="Server"
        Text='<%# Bind("ClockOUt") %>'
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
   <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="TTimebox" Text='<%# Eval("TotalTime") %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="TTimeEdit" Runat="Server"
        Text='Auto Generated' ReadOnly="true"
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Reasonbox" Text='<%# Eval("OutReason") %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="ReasonEdit" Runat="Server"
        Text='<%# Bind("OutReason") %>'
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="LocationBox" Text='<%# Eval("Location") %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="LocationEdit" Runat="Server"
        Text='<%# Bind("Location") %>'
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  <asp:TemplateField>
    <ItemTemplate>
      <asp:Label ID="Label15" Text='<%# Eval("LastName") %>' Runat="Server"
        Width="85"/>
    </ItemTemplate>
    <EditItemTemplate>
      <asp:TextBox id="EditPass" Runat="Server"
        Text='<%# Bind("LastName") %>'
        Width="85" Font-Size="8pt"/>
    </EditItemTemplate>
  </asp:TemplateField>
  
  </Columns>

<PagerStyle BackColor="#E0E0E0"></PagerStyle>

<EditRowStyle BackColor="Yellow"></EditRowStyle>

</asp:GridView>

        
        <br />
    </div>
    
    <p>
        <asp:Button ID="Button11" runat="server" OnClick="Show_All" Text="Show All" />
        <asp:Label ID="label" runat="server"></asp:Label>
    </p>
    
    </form>
</body>
</html>
