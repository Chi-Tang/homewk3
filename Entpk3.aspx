 
 
<Html>
<Body bgcolor="White">
<h2 align="center">@�}�v���g�R���{��@<h5 style="color: red">(�ήѥ���½�T�����o�T�ƿ�J)</h5></h2>
<h3 align="center">�R�������߸�,���M�w�R����,�Ƥ@����,�~���y���f,��e�b��,</h3>
<h3 align="center">����X�x����,�q���ҨD���Ʀp�U�C�M(����½�ѤT��):</h3>
<h3 align="center">�[�@�����Ħb�W,�̤l�f�f�f,�ߤ���........,�ôb�x��</h3>
<h3 align="center">�� ���ķO�d���̤l�H���g���H�ӫ��ܦN����V.</h3>

<H3>�п�J�򯫵��ĩҽ示��'�~��'�ܤ��T�ƽX<Hr></H3>

<Form runat="server">
�����ƽX�G<asp:TextBox id="Name" runat="server" /><p>
�~���ƽX�G<asp:TextBox id="Tel" runat="server"  /><p>
�ܤ��ƽX�G<asp:TextBox id="Addr" runat="server" /><asp:Button runat="server" Text=" ��J " OnClick="Button_Click" /><p>
<HR>
<asp:Label runat="server" id="Msg" ForeColor="Red" />
 </Form>
 
 
</Body>
</Html>

 <script Language="VB" runat="server">
 

   Sub Button_Click(sender As Object, e As EventArgs) 
      Dim URL
      URL = "Excpk3.aspx" & _
            "?Name="  & Server.URLEncode(Name.Text) & _
            "&Tel=" & Server.URLEncode(Tel.Text) & _
            "&Addr="  & Server.URLEncode(Addr.Text)
      Server.Transfer( URL )
   End Sub

</script>

 
