 
 
<Html>
<Body bgcolor="White">
<h2 align="center">@徐師易經卜卦程式@<h5 style="color: red">(用書本任翻三次取得三數輸入)</h5></h2>
<h3 align="center">卜卦首重心誠,先尋安靜場所,備一本書,洗手臉漱口,整容淨身,</h3>
<h3 align="center">雙手合掌閉目,默念所求之事如下七遍(完後翻書三次):</h3>
<h3 align="center">觀世音菩薩在上,弟子口口口,心中有........,疑惑困難</h3>
<h3 align="center">請 菩薩慈悲為弟子以易經卦象來指示吉凶方向.</h3>

<H3>請輸入佛神菩薩所賜內卦'外卦'變爻三數碼<Hr></H3>

<Form runat="server">
內卦數碼：<asp:TextBox id="Name" runat="server" /><p>
外卦數碼：<asp:TextBox id="Tel" runat="server"  /><p>
變爻數碼：<asp:TextBox id="Addr" runat="server" /><asp:Button runat="server" Text=" 輸入 " OnClick="Button_Click" /><p>
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

 
