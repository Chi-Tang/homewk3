<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

<html>
<body background="../B01.jpg" bgcolor="#FFFFFF">
<Form runat="server">
<h2>�ѤF�b���αK�X�H<hr></h2>
<blockquote>
Email�G<asp:TextBox runat="server" size="40" id="Email" />
   <asp:RequiredFieldValidator runat="server" Text="(���n���)"
        ControlToValidate="Email" EnableClientScript="False"
        Display="Dynamic" />
   <asp:RegularExpressionValidator runat="server"
        ControlToValidate="Email" Text="(Email ���t�� @ �Ÿ�)"
        ValidationExpression=".{1,}@.{3,}" 
        EnableClientScript="False" Display="Dynamic"/><Br>
�@<asp:Button runat="server" Text="�жǵ��ڱb���αK�X" OnClick="Forget_Click" /><br>
<Font Size=-1 Color=Blue>�п�J�z�� Email�A�M����U�u�жǵ��ڱb���αK�X�v</Font>
</blockquote>
<HR>
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</Form>
</body>
</html>

<script Language="VB" runat="server">

   Sub Forget_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
      If IsValid Then
         QueryDataAndSendTo()
      End If       
   End Sub

   Sub QueryDataAndSendTo()
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL As String, Body As String

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
       ''Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
       'Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
       'Dim Database = "Data Source=" & Server.MapPath( "/Hmath/UsersPwd.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
      Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()

      ' �ˬd Email �O�_�s�b
      SQL = "Select * From Users Where Email='" & Email.Text & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
      If Rd.Read() Then ' ��ܦ��@ Email �s�b
         Dim mail As MailMessage = New MailMessage

         mail.Subject = "�z���|�����"
         mail.To = Rd.Item("Email")
         mail.From = "tech.t1206@msa.hinet.net"   ' �令�t�κ��@�̪� e-mail
         mail.BodyFormat = MailFormat.Text 
         Body = "�ϥΪ̦W�١G" & Rd.Item("UserID") & vbCrLf
         Body = Body & "�@�@�@�K�X�G" & Rd.Item("Password") & vbCrLf 
         Body = Body & "�@�@�@�m�W�G" & Rd.Item("Name") & vbCrLf & vbCrLf
         Body = Body & "ASP.NET �����s�@�Х� �q�W"
         mail.Body = Body

         On Error Resume Next
           
           SmtpMail.SmtpServer = "msa.hinet.net"
          ''SmtpMail.SmtpServer = "smtp.gmail.com"
         'SmtpMail.SmtpServer = "seed.net.tw"
          SmtpMail.Send(mail) 

         If Err.Number <> 0 Then
            Msg.Text = Err.Description & "link or change SmtpServer"
         Else
            Msg.Text = "�b���αK�X�w�g�e�X�A���ˬd�z���H�c!"
         End If
      Else
         Msg.Text = "���@ Email �å��ӽбb��!"
      End If  
      Conn.Close()
   End Sub

</script>