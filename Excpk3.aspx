 
 <%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

 <%
  Dim TNo1 ,TNo2, TNo3,No1 ,No2, No3

   ''No = Request("No")
  'No1 = Request("No1")
  'No2 = Request("No2")
  'No3 = Request("No3")
  'TNo1 = Request("Name.Text")
  'TNo2 = Request("Tel.Text")
  'TNo3 = Request("Addr.Text")
 TNo1 = Request("Name")
 TNo2 = Request("Tel")
 TNo3 = Request("Add")
  No1 = Cint(TNo1)
  No2 = Cint(TNo2)
  No3 = Cint(TNo3)


 ''<script Language="VB" runat="server">

 '' Sub Button_Click(sender As Object, e As EventArgs) 
      '' Msg.Text = ""
      '' Dim URL 
      'No1 = Request("NNo1")
   'No2 = Request("NNo2")
   'No3 = Request("NNo3")
   ''
 
 Dim ZNAM()={"坤","艮","坎","巽","震","離","兌","乾","坤"}
 Dim ZNUM()={"0,0,0","0,0,1","0,1,0","0,1,1","1,0,0","1,0,1","1,1,0","1,1,1","0,0,0"}
 Dim ZNAMM()={"地","山","水","風","雷","火","澤","天","地"}
  'Dim No1 ,No2, No3
 Dim i ,j ,k ,n As integer
 Dim CGNo1,CGNo2,MGN1 ,MGN11,MGN2 ,MGN22,MDN3 ,MDN33   
 Dim GNam1,GNam2,MGNam1,MGNam2 
 Dim GNum1,GNum2
 Dim GNum1s,GNum2s
 Dim No12 
 Dim FGANK,FGCNK 
 Dim FGAN(5),FGCN(5)
 Dim FGAH(5),FGCH(5)
 Dim FGACK , FGCCK
  
 CGNo1= 8-(No1 mod 8)
 GNam1= ZNAMM(CGNo1)
 GNum1= ZNUM(CGNo1)
  GNum1s=SPLIT(GNum1, ",") 
  No12=Int(No1)+Int(No2)
CGNo2= 8-(No12 mod 8)
 GNam2= ZNAMM(CGNo2)
 GNum2= ZNUM(CGNo2)
 GNum2s=SPLIT(GNum2, ",") 
 ''FGAN(5)=GNum1s+GNum2s
  For i=0 to 2
    FGAN(i)=GNum1s(i)
    FGAN(i+3)=GNum2s(i)
  Next

  For j=0 to 5
    FGANK=Trim(FGAN(j))
      Select Case FGANK
      Case "1"
        FGAH(j)="一 "
      Case "0"
        FGAH(j)="- -"
      End  Select 
    
   FGCN(j)=FGAN(j)
   Next
   
   MDN3=Int((NO3 mod 6))
     if MDN3=0 then
        MDN33=5
      else
        MDN33=MDN3-1
     end if  

  For k=0 to 5
   if k=Int(MDN33) then 
     FGANK=Trim(FGAN(k))
     Select Case FGANK
      Case "1"
         FGCN(k)="0 "
      Case "0"
         FGCN(k)="1"
     End  Select  
   else
    FGCN(k)=FGAN(k) 
   end if  
  Next
 For n=0 to 5
    FGCNK=Trim(FGCN(n))
     Select Case FGCNK
      Case "1"
        FGCH(n)="一 "
      Case "0"
        FGCH(n)="- -"
     End  Select   
 NEXT
   '''Response.Write   "<TD>"&"卦序號=8-(卦碼號)"& MGN2 & MGN1  & "</TD>"
     MGN2=4*FGCN(3)+2*FGCN(4)+1*FGCN(5) 
           'CGCN22= 8-(CGCN2 mod 8)
            MGN22= (MGN2 mod 8)
           MGNam2= ZNAMM(MGN22)

      MGN1=4*FGCN(0)+2*FGCN(1)+1*FGCN(2) 
          ' CGCN11= 8-(CGCN1 mod 8)
           MGN11= (MGN1 mod 8)
           MGNam1= ZNAMM(MGN11)

     '' FGACK=PKCOD(GNam1,GNam2)
     '' FGCCK=PKCOD(MGNam1,MGNam2)
      
       FGACK= QueryDataAndSendTo(GNam1,GNam2)
       FGCCK= QueryDataAndSendTo(MGNam1,MGNam2)

     
     '  Response.Write (MGNam2)
     '  Response.Write (MGNam1)
     '   Response.Write (FGACK)
       
 Dim Row0 ="<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >"&"<TR>"&"<TD width=30>卦名"&"</TD>"&"<TD width=30>成卦</TD>"&"<TD width =30>變爻</TD>"&"<TD width =30>卦名</TD>"&"<TD width =30>之卦</TD>"&"</TR>"
 Dim Row1 = "<TR >" & "<TD width=30  RowSpan=3 ColSpan=1>" & GNam2 & "</TD>"& "<TD width=30>" & FGAH(5) & "</TD>"& "<TD width=30  >"& "</TD>"& "<TD width =30 RowSpan=3 ColSpan=1>" & MGNam2 &  "</TD>"& "<TD width =30>" &  FGCH(5) & "</TD>"& "</TR>"
 Dim Row2 = "<TR>" & "<TD width=30>" & FGAH(4) & "</TD>"& "<TD width=30  >"& "</TD>"& "<TD width =30>" & FGCH(4) & "</TD>"& "</TR>"
 Dim Row3 = "<TR>" & "<TD width=30>" & FGAH(3) & "</TD>"&  "<TD width=30>" & "</TD>"& "<TD width =30>" & FGCH(3) & "</TD>"&"</TR>"
 Dim Row4 = "<TR >" & "<TD width=30  RowSpan=3 ColSpan=1>" & GNam1 & "</TD>"& "<TD width=30>" & FGAH(2) & "</TD>"& "<TD width=30  >"& "</TD>"& "<TD width =30 RowSpan=3 ColSpan=1>" & MGNam1 &  "</TD>"& "<TD width =30>" &  FGCH(2) & "</TD>"& "</TR>"
 Dim Row5 = "<TR>" & "<TD width=30>" & FGAH(1) & "</TD>"& "<TD width=30  >"& "</TD>"& "<TD width =30>" & FGCH(1) & "</TD>"& "</TR>"
 Dim Row6 = "<TR>" & "<TD width=30>" & FGAH(0) & "</TD>"&  "<TD width=30>" & "</TD>"& "<TD width =30>" & FGCH(0) & "</TD>"&"</TR>"
 Dim Row7 = "<TR >"&"<TD width=76 >" & FGACK & "</TD>"&"<TD>" & "</TD>" & "<TD width=30  >"& MDN33 & "</TD>"&"<TD width=76>"& FGCCK &"</TD>"&"</TR>"&"</TABLE>"
 Dim Row8 = "<TR style='writing-mode:bt-rl' width =150>" & "尊重著作權:<u>免費分享</u>"& "</TR>"
 Dim Row9 ="<CENTER><H2><FONT color=#FF0000>"&"若需要<u>卦 盤 評 論 表</u>"&"</FONT><HR></H2>"
 Dim Row10 ="<CENTER><H2><FONT color=#FF0000>"&"    請 Email:  tech.t1206@gmail.com 即可聯絡,  謝謝"&"</FONT><HR></H2>"
 
 Dim RowT=Row0+Row1+Row2+Row3+Row4+Row5+Row6+Row7+Row8+Row9+Row10
 '' Response.Write (RowT)
 '' Response.Write (Row8)
  
  %>
 
  <HTML> 
  <style type="text/css">
   div {display: -ms-box;position:relative; top:40px+20px; width:130px+20px;
       writing-mode:tb-rl;background-color: red;column-count:4;
       -ms-grid-row: 4; }
  #content01 {position:absolute; top:40px; width:130px; heigtht:130px;
              background-color: red;-ms-grid-column:2; -ms-grid-row: 2; }
  </style>

 
 <BODY bgcolor="#FFFFFF">
  <CENTER><H2> 易 經 八 卦 盤 與 評 論 表 如 下<HR></H2>
  <TABLE Border=1 Width=60%  BGCOLOR=#FFFF00>
   <!--<TR><TD width=30>卦名</TD><TD width=30>成卦</TD><TD width =30>變爻</TD><TD width =30>卦名</TD><TD width =30>之卦</TD></TR>
   -->
  <%  
 ''Dim Row1 ,Row2 ,Row3 ,Row4 ,Row5 ,Row6 ,Row7 ,Row8 ,RowT
        Response.Write (RowT)
  %>

</TABLE></CENTER>
 </BODY>
 
  <!-- <CENTER><H2><FONT color=#FF0000>若需要<u>卦 盤 評 論 表</u></FONT><HR></H2>
 <CENTER><H2><FONT color=#FF0000>    請 Email:  tech.t1206@gmail.com 即可聯絡,  謝謝</FONT><HR></H2>
  -->
</Html>
 
 <script Language="VB" runat="server">
    Function  QueryDataAndSendTo(XNam1,XNam2)
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rd As OleDbDataReader, SQL As String, Body As String
      Dim mad11, mad22 As String
      Dim k As integer
       k=0
      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
       ''Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
       'Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
        'Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "/HSU-WN/UsersPwd.mdb" )
      Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()

      ' 檢查 Email 是否存在
      ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
      'SQL = "Select * From 八卦表 Where 內卦='" & Emad1.Text & "'"
       SQL = "Select * From 八卦表 "
      Cmd = New OleDbCommand( SQL, Conn )
      Rd = Cmd.ExecuteReader()
    While Rd.Read()
       If Rd.Item(0)=XNam1 Then ' 表示此一 Email 存在
        'If Rd.Item("內卦")=Emad1.Text Then ' 表示此一 Email 存在
         'If Rd.Read() Then ' 表示此一 Email 存在
         ' Dim mad1 As string
       for k = 1 to Rd.Fieldcount -1
        '' mad11 = Rd.Item("天")
          if Rd.GetName(k) = XNam2 then
             
             mad22 =Rd.Item(k)
       
          End If 

         ''mad11 =mad11+ Rd.Item(k)
        ''mad11 =mad11+ Rd.GetName(k)

       next 
      end if
  
  End While   
     Rd.close()
      Conn.Close()
    
     ''QueryDataAndSendTo=mad22
      
      ''Response.Write (mad11)
     Return mad22 
      
 End Function 
 
  </script>
 