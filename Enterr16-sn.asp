 
 <!-- #include virtual="/HSU-PK/DB.fun" -->
 <%
  ''<!-- #include file="Login-1.asp" -->
  %>
<%
mdbFile = "/HSU-fundb/UsersPwd.mdb"
 '' mdbFile = "/Hmath/UsersPwd.mdb"
 ''mdbFile = "UsersPwd.mdb"
 mdbPassword = "kj6688"

 MySelf = Request.ServerVariables("PATH_INFO")
   Lesson = Request("Lesson")
   No = Request("No")
   'Name = Request("Name")
  Sex=Request("Sex")
SECTM=Request("SECTM")
TNUM=Request("TNUM")
DNUM=Request("DNUM")
HNUM=Request("HNUM")
LYR=Request("LYR")

 'SECTMS=SPLIT(SECTM, ",")
 'TNUMS=SPLIT(TNUM, ",") 
 'CNm=Request("NUM")
ZK= Array("��","�A","��","�B","��","�v","��","��","��","��")
ZG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")

Mssg = "ok"
''On Error Resume Next 

If Request("Send") <> Empty Then
   SQL = "Select * From BIRTH " 
 Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
   ' SQL = "Select * From BIRTH " 
 'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
  ' SQL = "Select * From ���Z�� " 
  ' SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
  ' Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
  ''    SQLL = "Select * From "&Lesson&"k" 
  ''    Set rs = GetMdbRecordset( "Testac-1.mdb", SQLL )
     n=0
     TNum1=0
   
   ERNDSN  SECTM,TNUM,DNUM,HNUM
    ''RNDSN  SECTM,TNUM
    'For  k=0 to Ubound(SECTMS)
     ' RNDSN SECTMS(k),TNUMS(k)
    'NEXT
  YKKN=ERNDSN(SECTM, TNUM, DNUM, HNUM)
  YKKG=GRNDSN(SECTM, TNUM, DNUM, HNUM)
  LYGG=LRNDSN(LYR, "2", "20")
  ONLER= YERNDSN(SECTM, TNUM, DNUM, HNUM)

  ''LNDAT=DATCld(2012,6,6)
  
  Response.Write "<TR><TD>�U�~�z:"& ONLER & "</TD></TR>"
  
  '' On Error Resume Next 
  
  'Set conn = GetMdbConnection("Test1.mdb")
  Set cmd = Server.CreateObject( "ADODB.Command" )
  Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" �@
   '    SQLD2 ="Delete From "&Lesson&Trim(Cstr(No))&"A"
 ' cmd.CommandText = SQLD2
'  cmd.Execute

     
  '' On Error Resume Next 
  ' If Err.Number = 0 Then 
  '    Response.Write Mssg
  '  Else
   '   Response.Write Err.Number
  '  End If
       
  'SQL1 ="Insert into ASP Select * FROM "&Lesson&"K"&"Where �аO=+1"
  ' SQL1 ="Insert into "&Lesson&Trim(Cstr(No))& " Select * From "&Lesson&"K"&" Where �аO=100"
 '  cmd.CommandText = SQL1
  '   cmd.Execute
 '  SQLU ="Update "&Lesson&"K"&" Set �аO=-1"  
  '  cmd.CommandText = SQLU
  '   cmd.Execute
  '' Response.Redirect "TestKac-1t.asp?" & Request.QueryString
 ''  Response.Redirect "Excer2-sn.asp?" & "YKK=" & YKKN & "&" & "YKG=" & YKKG & "&" & "LYG=" & LYGG & "&" & Request.QueryString
  ''  Response.Redirect "Excer2-1.asp?" & "YKK=" & YKKN & "&" & "YKG=" & YKKG & "&" & "LYG=" & LYGG & "&" & Request.QueryString
  
 %>

 
 <%   
 
End If    

 %> 
 
 <html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>�}�Ѯv���L��ƾǶ�a</title>
</head>

<body  onload=initialize() background="../b01.jpg">
 <%
  Dim SLY,SLM,TLD
 FUNCTION YERNDSN(SECTM, TNUM, DNUM, HNUM)
  
 ''SYY= CLD.SECTM.selectedIndex+1912
 ''SMM= CLD.TNUM.selectedIndex
 '' TDD= CLD.DNUM.selectedIndex+1
 SLY=SECTM
 SLM=TNUM
 TLD=DNUM 
 Dim  Driver , DBpath,Param , LANDA, LDBPH
 Dim conn, rsk
  Driver = "Driver={Microsoft Excel Driver (*.xls)};"
  ''DBPath = "DBQ=" & Server.MapPath("Excel02.xls""TEST201111.xls")
  DBPath = "DBQ=" & Server.MapPath("TEST"&SLY&SLM&TLD&".XLS")
    '' "C:\\Inetpub\\wwwroot\\Hsu-pk\\TEST"+SYY+(SMM+1)+TDD+".XLS"
   LDBPH ="TEST"&SLY&SLM&TLD&".XLS"
  'DBPath = "DBQ=" & Server.MapPath("TEST"&SLY&SLM&TLD&".XLS")
   'DBPath = "DBQ=" & Server.MapPath("TEST201308.xls")
   Param =Driver & "ReadOnly=0;" & DBPath
   SQL = "Select * From [A1:I30]"
   ''SQL = "Select * From ��q��"
    '  Set GetExcelConnection = GetConnection(Driver & "ReadOnly=0;" & DBPath)
    ' Dim conn
    ' On Error Resume Next  ''If Err.Number <> 0 Then Exit Function
    ' Set GetConnection = Nothing  
   Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Param
   Set rsk = Server.CreateObject("ADODB.Recordset")
   rsk.Open SQL, conn, 2, 2

  ' Part I�G��X�u���Y�W�١v
  For i=0 to rsk.Fields.Count-1
     LANDA=LANDA + rsk(i).Name 
  Next
 '' Response.Write "<TD>" & rsk(i).Name & "</TD>"
 '' Response.Write LANDA
  ''YERNDSN = LANDA
   YERNDSN = LDBPH&LANDA
 END FUNCTION 
%>
 
<%  '�p��~���z��
 FUNCTION ERNDSN(SECTM, TNUM, DNUM, HNUM)

' SUB ERNDSN(SECTM, TNUM, DNUM, HNUM)
  '  SQL = "Select * From BIRTH " 
 'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
     ' TNUM1=30 
    ' TNUM1=TNUM1+TNUM 
     '  SECTM=SECTM+0
   D1=  #1912/2/18#
   D2=DateSerial(SECTM,TNUM,DNUM)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
 '/ Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
 '/  Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
 ' Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 '/ Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
    rs.AddNew
  'rs("�Ǹ�") = CLng(UserID)
  'rs("�Ǹ�") = CLng(Trim(Right(UserID,6)))
  'rs("�Ǹ�") = Trim(Right(UserID,6))
 rs("�Ǹ�") = Trim(Right(NO,12))
 rs("�m�W") =  Name
 rs("�~") = SECTM
 rs("��") = TNUM
 rs("��") = DNUM 
 rs("��") = HNUM
 rs("YK") = ZK(YK)
 rs("YG") = ZG(YG)
 rs("MG") = ZG(MG)
 rs("DK") = ZK(DK)
 rs("DG") = ZG(DG)
 rs("HG") = HNUM
 rs.Update

   ERNDSN=ZK(YK)
 END FUNCTION 
' END SUB
  %>    
<%  '�p��~��
 FUNCTION GRNDSN(SECTM, TNUM, DNUM, HNUM)

   D1=  #1912/2/18#
   D2=DateSerial(SECTM,TNUM,DNUM)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
 '/ Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
 '/ Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
'/ Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
   
 
 GRNDSN=ZG(YG)

 END FUNCTION 

  %>    
<%  '�p��y�~��
 FUNCTION LRNDSN(LYR, MON, DT)

   D1=  #1912/2/18#
   D2=DateSerial(LYR,MON, DT)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
 '/ Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
 '/ Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
'/ Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
   
 
 LRNDSN=ZG(YG)

 END FUNCTION 

  %>    


<h2 align="center">�}�v�Ч@���L��Ƶ{��</h2>

<hr>

<h2>��ܩR�z���:</h2>
<body>
<blockquote>
   <form  onsubmit=SDATCld() name= CLD action="<%=Myself%>" method="GET">
    <!-- <form onsubmit=SDATCld() name= CLD action= Excer2-1.asp  method="GET">-->
    
       <p>��ءG<select name="Lesson" size="1"> 
            <option IsSelected("ZEWU", Lesson)>ZEWU</option> 
        </select></p> 
        <p>�п�J���N�Ʀr�G<input type="text" size="20" name="No" Value="<%=No%>"></p>     
     <!---   �m�W�G<input type="text" size="20" name="Name"  Value="<%=Name%>"> --> 
    
        <p>�ʧO�G<select name="Sex" size="1">                                                                                                                                                                                                                                                                                                                 
            <option value="�k">�k</option><option value="�k">�k</option> 
             </select>
          </p>

           
      <h5> �A��褸?�~ �@ &���              &�鸹            &�ɨ�                      &���~�褸?�~ :</h5>                                                                                                                                                                                                               
         <p>�G<select name="SECTM" size="6"> 
            <option value="1912">1912</option><option value="1913">1913</option><option value="1914">1914</option><option value="1915">1915</option> 
            <option value="1916">1916</option><option value="1917">1917</option><option value="1918">1918</option><option value="1919">1919</option> 
             <option value="1920">1920</option><option value="1921">1921</option><option value="1922">1922</option><option value="1923">1923</option>
             <option value="1924">1924</option><option value="1925">1925</option><option value="1926">1926</option><option value="1927">1927</option>
             <option value="1928">1928</option><option value="1929">1929</option><option value="1930">1930</option><option value="1931">1931</option>
             <option value="1932">1932</option><option value="1933">1933</option><option value="1934">1934</option><option value="1935">1935</option>
             <option value="1936">1936</option><option value="1937">1937</option><option value="1938">1938</option><option value="1939">1939</option>
             <option value="1940">1940</option><option value="1941">1941</option><option value="1942">1942</option><option value="1943">1943</option>
             <option value="1944">1944</option><option value="1945">1945</option><option value="1946">1946</option><option value="1947">1947</option>
             <option value="1948">1948</option><option value="1949">1949</option><option value="1950">1950</option><option value="1951">1951</option>
             <option value="1952">1952</option><option value="1953">1953</option><option value="1954">1954</option><option value="1955">1955</option>
             <option value="1956">1956</option><option value="1957">1957</option><option value="1958">1958</option><option value="1959">1959</option>
             <option value="1960">1960</option><option value="1961">1961</option><option value="1962">1962</option><option value="1963">1963</option>
             <option value="1964">1964</option><option value="1965">1965</option><option value="1966">1966</option><option value="1967">1967</option>
             <option value="1968">1968</option><option value="1969">1969</option><option value="1970">1970</option><option value="1971">1971</option>
             <option value="1972">1972</option><option value="1973">1973</option><option value="1974">1974</option><option value="1975">1975</option>
             <option value="1976">1976</option><option value="1977">1977</option><option value="1978">1978</option><option value="1979">1979</option>
             <option value="1980">1980</option><option value="1981">1981</option><option value="1982">1982</option><option value="1983">1983</option>
             <option value="1984">1984</option><option value="1985">1985</option><option value="1986">1986</option><option value="1987">1987</option>
             <option value="1988">1988</option><option value="1989">1989</option><option value="1990">1990</option><option value="1991">1991</option>
             <option value="1992">1992</option><option value="1993">1993</option><option value="1994">1994</option><option value="1995">1995</option>
             <option value="1996">1996</option><option value="1997">1997</option><option value="1998">1998</option><option value="1999">1999</option>
             <option value="2000">2000</option><option value="2001">2001</option><option value="2002">2002</option><option value="2003">2003</option>
             <option value="2004">2004</option><option value="2005">2005</option><option value="2006">2006</option><option value="2007">2007</option>
             <option value="2008">2008</option><option value="2009">2009</option><option value="2010">2010</option><option value="2011">2011</option>
             <option value="2012">2012</option><option value="2013">2013</option><option value="2014">2014</option><option value="2015">2015</option>
                      
              </select> 
           <!.../p...>                                                                                                                                                                                                                                                                                                                                                                  
          �G<select name="TNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                  
            <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
             <option value="9">9</option><option value="10">10</option> 
             <option value="11">11</option><option value="12">12</option> 
            </select>
            <!.../p...>                                                                                                                                                                                                               
          �G<select name="DNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
             <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
             <option value="9">9</option><option value="10">10</option> 
             <option value="11">11</option><option value="12">12</option> 
             <option value="13">13</option><option value="14">14</option> 
             <option value="15">15</option><option value="16">16</option> 
             <option value="17">17</option><option value="18">18</option> 
             <option value="19">19</option><option value="20">20</option> 
             <option value="21">21</option><option value="22">22</option> 
             <option value="23">23</option><option value="24">24</option> 
             <option value="25">25</option><option value="26">26</option> 
             <option value="27">27</option><option value="28">28</option> 
             <option value="29">29</option><option value="30">30</option> 
             <option value="31">31</option> 
            </select>
            <!.../p...>                                                                                                                                                                                                               
          �G<select name="HNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
            <option value="�l">23~01</option><option value="��">01~03</option> 
             <option value="�G">03~05</option><option value="�f">05~07</option> 
             <option value="��">07~09</option><option value="�w">09~11</option> 
             <option value="��">11~13</option><option value="��">13~15</option> 
             <option value="��">15~17</option><option value="��">17~19</option> 
             <option value="��">19~21</option><option value="��">21~23</option> 
            </select>
                                                                                                                                                                               
         �G<select name="LYR" size="6">                           
             <option value="2007">2007</option><option value="2008">2008</option>  <option value="2009">2009</option><option value="2010">2010</option>
              <option value="2011">2011</option><option value="2012">2012</option><option value="2013">2013</option><option value="2014">2014</option>
             <option value="2015">2015</option><option value="2016">2016</option><option value="20176">2017</option><option value="2018">2018</option>
             <option value="2019">2019</option>  <option value="2020">2020</option><option value="2021">2021</option>
                <!.../p...>     
            </select>                            
          </p> 
         <p><input type="submit" Name="Send" value="�w��i�J��a"> </p> 
         <p>�@ </p> 
         
         
    </form> 
  </blockquote>  
 
<hr> 
<FONT Color=Red><%=Msg%></FONT> 
<center><a href="http://www.ineedhits.com/free-tools/submit-free.aspx?source=FTSFbutton"><img src="http://www.ineedhits.com/images/trackingbuttons/SFbutton.gif?ref=1563375" border="0" height="32" width="90" alt="Submit your website to 20 Search Engines - FREE with ineedhits!"></a></center>
<center><a href="http://www.ineedhits.com/optimization/optimization.aspx" style="font-family: Arial; font-size:11px; color: Gray; font-weight:bolder; orientation:Portrait; text-decoration:none">SEO Services</a></center>

</boody> 
<!-- <script>
    document.onselectionchange=__OnSelectionChange;
       var running=false;
     function __OnSelectionChange()
       { 
       if (running==true) return;
          running=true;
       document.selection.empty();
       running=false;       
        }
  </script>--> 
   
 
  <script language="JavaScript">
  
  var LunData=new Array(
"0A4D0","0D250","1D295","0B550","056A0","0ADA2","095B0","14977","049B0","0A4B0",
"0B4B5","06A50","06D40","1AB54","02B60","09570","052F2","04970","06566","0D4A0",
"0EA50","16A95","05AD0","02B60","186E3","092E0","1C8D7","0C950","0D4A0","1D8A6",
"0B550","056A0","1A5B4","025D0","092D0","0D2B2","0A950","0B557","06CA0","0B550",
"15355","04DA0","0A5B0","14573","052B0","0A9A8","0E950","06AA0","0AEA6","0AB50",
"04B60","0AAE4","0A570","05260","0F263","0D950","05B57","056A0","096D0","04DD5",
"04AD0","0A4D0","0D4D4","0D250","0D558","0B540","0B6A0","195A6","095B0","049B0",
"0A974","0A4B0","0B27A","06A50","06D40","0AF46","0AB60","09570","04AF5","04970",
"064B0","074A3","0EA50","06B58","05AC0","0AB60","096D5","092E0","0C960","0D954",
"0D4A0","0DA50","07552","056A0","0ABB7","025D0","092D0","0CAB5","0A950","0B4A0",
"0BAA4","0AD50","055D9","04BA0","0A5B0","15176","052B0","06930","07954","06AA0",
"0AD50","05B52","04B60","0A6E6","0A4E0","0D260","0EA65","0D520","0DAA0","076A3",
"096D0","04AFB","04AD0","0A4D0","1D0B6","0D250","0D520","0DD45","0B5A0","056D0",
"055B2","049B0","0A577","0A4B0","0AA50","1B255","06D20","0ADA0","14B63","09370",
"049F8","04970","064B0","168A6","0EA50","06B20","1A6C4","0AAE0","092E0","0D2E3",
"0C960","0D557","0D4A0","0DA50","05D55","056A0","0A6D0","055D4","052D0","0A9B8");
var Today = new Date();
var Ny = Today.getFullYear();
var Nm = Today.getMonth();
var Nd = Today.getDate();
var cld,Selday;
var NHoliday = new Array(
"0101*����",
"0111 �q�k�`",
"0115 �Įv�`",
"0123 �ۥѤ�",
"0204 �A���`",
"0214 ���H�`",
"0215 ���@�`",
"0219 �s�ͬ��B�ʬ�����",
"0228*�M��������",
"0301 �L�и`",
"0305 ���l�x�`",
"0308 ���k�`",
"0312 �Ӿ�`",
"0317 ����`",
"0320 �l�F�`",
"0321 ��H�`",
"0325 ���N�`",
"0326 �s���`",
"0329 �C�~�`",
"0330 �X���`",
"0401 �M�H�`",
"0404 �����`",
"0405 ���ָ`",
"0407 �å͸`",
"0422 �@�ɦa�y��",
"0501*�Ұʸ`",
"0504 �����`",
"0505 �R�и`",
"0510 �]��`",
"0512 �@�h�`",
"0603 �T�ϸ`",
"0606 �u�{�v�`",
"0609 �K���`",
"0615 ĵ��`",
"0630 �|�p�v�`",
"0701 �����`",
"0711 ����`",
"0712 Ť�׸`",
"0808 ���˸`",
"0814 �ŭx�`",
"0827 �G���\�Ϩ�",
"0901 �O�̸`",
"0903 �x�H�`",
"0909 ��|�`",
"0913 �k�ߤ�",
"0928 �Юv�`",
"1006 �ѤH�`",
"1010*��y������",
"1021 �ع��`",
"1025 �x�W���_�`",
"1031 �U�t�`",
"1101 �ӤH�`",
"1111 �u�~�`",
"1117 �ۨӤ��`",
"1112 ����Ϩ���",
"1121 ���Ÿ`",
"1205 �����`",
"1210 �H�v�`",
"1212 �˧L�`",
"1225 ��ˬ�����",
"1227 �ؿv�v�`",
"1228 �q�H�`",
"1231 ���H�`");
var WHoliday = new Array(
"0520 ���˸`",
"0716 �X�@�`",
"0730 �Q���а�a�g",
"1144 �P���`");
var LHoliday = new Array(
"0101*�K�`",
"0102*�^�Q�a",
"0103*����",
"0104 �ﯫ",
"0105 �}��",
"0109 �Ѥ���",
"0115 ���d�`",
"0202 �Y��",
"0323 ������",
"0408 �D��`",
"0505*�ݤȸ`",
"0701 �}����",
"0707 �C�i���H�`",
"0715 �����`",
"0800 ������",
"0815*����`",
"0909 �����`",
"1208 þ�K�`",
"1216 ����",
"1224 �e��",
"0100*���i");

function initialize() {

 var Today = new Date();
var Ny = Today.getFullYear();
var Nm = Today.getMonth();
var Nd = Today.getDate();

// CLD.SY.selectedIndex=Ny-1912;
/// CLD.SM.selectedIndex=Nm;
  
}
var NMonthDays=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
function isLeapYear(y,m) {
 if(m==1)
    return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
 else
    return(NMonthDays[m]);
}

function lYearDays(y) {
var i,k, sum = 0; 
  k=StrToInt(y,5);
  for(i=1;i<13;i++) sum += (k & (0x10000>>i))? 30 : 29;
   return(sum+leapDays(y));
}

function leapDays(y) {
  if(leapMonth(y)) return( (StrToInt(y,1)&0xf)? 30: 29);
  else return(0);
}

function leapMonth(y) {
 return(StrToInt(y,5) & 0xf);
}

function LmonthDays(y,m) {
  return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );
}

function StrToInt(yx,vx){
 var sr;
  sr=LunData[yx-1912];
   sr=sr.substring(0, vx);
   return (parseInt('0x'+sr));
}

function Lunar(objDate) {
 var i, leap=0, temp=0;
 var offset = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),
                        objDate.getDate()) - Date.UTC(1912,1,18))/86400000;
 for(i=1912; i<2072 && offset>0; i++) { temp=lYearDays(i); offset-=temp; }
  if(offset<0) { offset+=temp; i--; }
   this.year = i;
  leap = leapMonth(i); 
  this.isLeap = false;
 for(i=1; i<13 && offset>0; i++) {
   if(leap>0 && i==(leap+1) && this.isLeap==false)
     { --i; this.isLeap = true; temp = leapDays(this.year); }
   else
    { temp = LmonthDays(this.year, i); }
     if(this.isLeap==true && i==(leap+1)) this.isLeap = false;
     offset -= temp;
 }

 if(offset==0 && leap>0 && i==leap+1)
    if(this.isLeap)
       { this.isLeap = false; }
    else
       { this.isLeap = true; --i; }
  if(offset<0){ offset += temp; --i; }
   this.month = i;
   this.day = offset + 1;
}
var Gan=new Array("��","�A","��","�B","��","�v","��","��","��","��");
var Zhi=new Array("�l","��","�G","�f","��","�x","��","��","��","��","��","��");
function GanZhi(num) {
 return(Gan[num%10]+Zhi[num%12]);
}


function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,jMonth,cDay) {
  this.isToday    = false;
  this.sYear      = sYear;   
   this.sMonth     = sMonth; 
    this.sDay       = sDay;  
     this.week       = week; 
  this.lYear      = lYear;   
   this.lMonth     = lMonth; 
    this.lDay       = lDay;  
     this.isLeap     = isLeap; 
  this.cYear      = cYear;   
   this.cMonth     = cMonth; 
    this.jMonth     = jMonth;
     this.cDay       = cDay; 
     this.color      = '';
  this.Lholiday = ''; 
   this.Nholiday = ''; 
    this.solarTerms    = '';
}
var TermData = new Array(0,21324,42505,63868,85407,107110,128977,151002,173218,195611,218134,240768,263418,286062,308631,331096,353423,375568,397546,419292,440895,462344,483626,504891);
function sTerm(y,n) {
 var offDate = new Date( ( 31556925974.7*(y-1912) + TermData[n]*60000  ) + Date.UTC(1912,0,7,0,7) );
 return(offDate.getUTCDate());
}
var SolarTerm = new Array("�p�H","�j�H","�߬K","�B��","���h","�K��","�M��","�\�B","�߮L","�p��","�~��","�L��","�p��","�j��","�߬�","�B��","���S","���","�H�S","����","�ߥV","�p��","�j��","�V��");
var dStr1 = new Array('��','�@','�G','�T','�|','��','��','�C','�K','�E','�Q');
function calendar(y,m) {
 var sDObj, lDObj, Ly, Lm, Ld=1, Lp, Lx=0, tmp1, tmp2, tmp3;
 var Cy, Cm, Jm, Cd; 
 var Ldpos = new Array(3);
 var n = 0;
 var firstLM = 0;
 sDObj = new Date(y,m,1,0,0,0,0); 
 this.length    = isLeapYear(y,m);
 this.firstWeek = sDObj.getDay(); 
  if(m<2)  Cy=GanZhi(y-1912+47); 
 else Cy=GanZhi(y-1912+48);
 var term2=sTerm(y,2);
 var firstNode = sTerm(y,m*2);
 Jm = GanZhi((y-1912)*12+m+36);
 var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+21185+12;
 
 for(var i=0;i<this.length;i++) {
    if(Ld>Lx) {
       sDObj = new Date(y,m,i+1);
       lDObj = new Lunar(sDObj);
       Ly    = lDObj.year;
       Lm    = lDObj.month;
       Ld    = lDObj.day;
       Lp    = lDObj.isLeap;
       Lx    = Lp? leapDays(Ly): LmonthDays(Ly,Lm);
   
      if(n==0)  firstLM = Lm;
       Ldpos[n++] = i-Ld+1;
    }

    if(m==1 && (i+1)==term2) Cy=GanZhi(y-1912+48);
     if((i+1)==firstNode) Jm = GanZhi((y-1912)*12+m+37);
	Cm= GanZhi((Ly-1912)*12+Lm+37);
    Cd = GanZhi(dayCyclical+i);
    this[i] = new calElement(y,m+1,i+1,dStr1[(i+this.firstWeek)%7],
	                        Ly,Lm,Ld++,Lp,Cy,Cm,Jm,Cd);
   }

  tmp1=sTerm(y,m*2  )-1;
   tmp2=sTerm(y,m*2+1)-1;
 this[tmp1].solarTerms = SolarTerm[m*2];
  this[tmp2].solarTerms = SolarTerm[m*2+1];
   if(m==3) this[tmp1].color = 'red';

 for(i in NHoliday)
  if(NHoliday[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
   if(Number(RegExp.$1)==(m+1)) {
    if(Number(RegExp.$2)<=this.length){
     this[Number(RegExp.$2)-1].Nholiday += RegExp.$4 + ' ';
      if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red';
       }
     }
 for(i in WHoliday)
  if(WHoliday[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
   if(Number(RegExp.$1)==(m+1)) {
    tmp1=Number(RegExp.$2);
     tmp2=Number(RegExp.$3);
      if(tmp1<5)
       this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].Nholiday += RegExp.$5 + ' ';
      else {
       tmp1 -= 5;
        tmp3 = (this.firstWeek+this.length-1)%7;
         this[this.length - tmp3 - 7*tmp1 + tmp2 - (tmp2>tmp3?7:0) - 1 ].Nholiday += RegExp.$5 + ' ';
         }
       }
 for(i in LHoliday)
  if(LHoliday[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
    tmp1=Number(RegExp.$1)-firstLM;
     if(tmp1==-11) tmp1=1;
      if(tmp1 >=0 && tmp1<n) {
       tmp2 = Ldpos[tmp1] + Number(RegExp.$2) -1;
        if( tmp2 >= 0 && tmp2<this.length && this[tmp2].isLeap!=true) {
         this[tmp2].Lholiday += RegExp.$4 + ' ';
          if(RegExp.$3=='*') this[tmp2].color = 'red';
         }
       }
    }
 if(m==2 || m==3) {
  var estDay = new easter(y);
   if(m == estDay.m)
    this[estDay.d-1].Nholiday = this[estDay.d-1].Nholiday+'�_���`';
 }
 if((this.firstWeek+12)%7==5) this[12].Nholiday += '�¦�P����';
 if(y==Ny && m==Nm) {
  this[Nd-1].isToday = true;
   this[Nd-1].color ='#ff00ff';
 }
}

//////////////////////////////////////////////////////////
  
  var dStr2 = new Array('��','�Q','��','��','�m');
function cDay(d){
 var s;
 switch (d) {
    case 10:
       s = '��Q'; break;
    case 20:
       s = '�G�Q'; break;
    case 30:
       s = '�T�Q'; break;
    default :
       s = dStr2[Math.floor(d/10)];
       s += dStr1[d%10];
 }
 return(s);
}

 function GetZhi(s){
   s = s.substr(1, 1);
  for(i=0;i<12;i++)
   if(s==Zhi[i]) return(i);
}	

function mCls() {
 for(i=0;i<42;i++) {
    mObj=eval('SD'+ i);
	 mObj.className='';
  }	 
 }  

 </script>
  <script language="JavaScript">
 //////////ctiveXObject ����/////////////////////////////////////
 var sk,st11,st22, cldt;
 var SYY,SMM,TDD;
 
 function SDATCld(SYY,SMM,TDD) {
  var i,sD,st,st1,st2,st3,size,Lastday;
 var p1,p2 ;
 SYY= CLD.SECTM.selectedIndex+1912;
 SMM= CLD.TNUM.selectedIndex;
 TDD= CLD.DNUM.selectedIndex+1;
   cldt = new calendar(SYY,SMM);
   //cldt = new calendar(2001,1);
  // cldt1 = new calendar(2001,1);
  // cldt2 = new calendar(2001,2);
   //cldt2[] =cldt2[].concat(cldt1[]);
   //st3 =cldt1[].concat(cldt2[]);
   //var CADObject;
  // CADObject = GetObject("C:\\CAD\\SCHEMA.CAD");

  var Today = new Date();
   //var Nyy = Date.UTC(Today);
   //var Nyy = Date.UTC(Today.getFullYear());
   var Nyy = Today.getFullYear();
   var Nmm = Today.getMonth();
   var Ndd = Today.getDate();

  var ExcelSheet;
   ExcelApp = new ActiveXObject("Excel.Application");
   ExcelSheet = new ActiveXObject("Excel.Sheet");
   
    // �z�L Application �������� Excel�C
   ExcelSheet.Application.Visible = true; 
  //for (var j=0 ;j<cldt.length;j++){ 
      //sD = i - cldt.firstWeek;
    sD = 10 ;
   // sD = TDD ;
   //  p1 = cldt.length;
    //  p2 =cldt.length+cldt1.length;
   
 //var xlBook = xls.Workbooks.Add; 
//var xlsheet = xlBook.Worksheets(1); 
 if(sD>-1 && sD<cldt.length) {
 
 for (var j=0;j<cldt.length;j++) {
 //xlsheet.Cells(i+1,j+1).value = objTable.rows[i].cells[j].innerHTML; 
 // �b�u�@���Ĥ@���x�s�椤��J�@�Ǥ�r�C					
 //ExcelSheet.ActiveSheet.Cells(1,1).Value = "Thih is column,row1";	
  ExcelSheet.ActiveSheet.Cells(j+1,1).Value =cldt[j].sYear;
  ExcelSheet.ActiveSheet.Cells(j+1,2).Value = cldt[j].sMonth;
  ExcelSheet.ActiveSheet.Cells(j+1,3).Value = cldt[j].sDay;
  ExcelSheet.ActiveSheet.Cells(j+1,4).Value =cldt[j].lYear;
  ExcelSheet.ActiveSheet.Cells(j+1,5).Value = cldt[j].lMonth;
  ExcelSheet.ActiveSheet.Cells(j+1,6).Value = cldt[j].lDay;
  ExcelSheet.ActiveSheet.Cells(j+1,7).Value =cldt[j].cYear;
  ExcelSheet.ActiveSheet.Cells(j+1,8).Value = cldt[j].cMonth;
  ExcelSheet.ActiveSheet.Cells(j+1,9).Value = cldt[j].cDay;
//   }
// for (var j=cldt.length;j<cldt1.length;j++) {
 //xlsheet.Cells(i+1,j+1).value = objTable.rows[i].cells[j].innerHTML; 
// �b�u�@���Ĥ@���x�s�椤��J�@�Ǥ�r�C					
 //ExcelSheet.ActiveSheet.Cells(1,1).Value = "Thih is column,row1";	
 // ExcelSheet.ActiveSheet.Cells(p1+j+1,1).Value =cldt1[j].sYear;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,2).Value = cldt1[j].sMonth;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,3).Value = cldt1[j].sDay;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,4).Value =cldt1[j].lYear;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,5).Value = cldt1[j].lMonth;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,6).Value = cldt1[j].lDay;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,7).Value =cldt1[j].cYear;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,8).Value = cldt1[j].cMonth;
//  ExcelSheet.ActiveSheet.Cells(p1+j+1,9).Value = cldt1[j].cDay;
 /// �b�u�@���Ĥ@���x�s�椤��J�@�Ǥ�r�C					
 //ExcelSheet.ActiveSheet.Cells(1,1).Value = "Thih is column,row1";	
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,1).Value =cldt2[j].sYear;
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,2).Value = cldt2[j].sMonth;
 // ExcelSheet.ActiveSheet.Cells(p2+j+1,3).Value = cldt2[j].sDay;
 // ExcelSheet.ActiveSheet.Cells(p2+j+1,4).Value =cldt2[j].lYear;
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,5).Value = cldt2[j].lMonth;
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,6).Value = cldt2[j].lDay;
 // ExcelSheet.ActiveSheet.Cells(p2+j+1,7).Value =cldt2[j].cYear;
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,8).Value = cldt2[j].cMonth;
//  ExcelSheet.ActiveSheet.Cells(p2+j+1,9).Value = cldt2[j].cDay;

 
 }

 
 /// �x�s�Ӥu�@��C
 //ExcelSheet.SaveAs("C:\\Inetpub\\wwwroot\\Hsu-pk\\TEST"+Nyy+Nmm+Ndd+".XLS");	
  ExcelSheet.SaveAs("C:\\Inetpub\\wwwroot\\Hsu-pk\\TEST"+SYY+(SMM+1)+TDD+".XLS");	
 //ExcelSheet.copy("C:\\TEST.XLS","After");	
  //ExcelSheet.SaveAs("C:\\TEST.XLS");	

    //// �ϥ� Application ���� Quit ��k������ Excel�C
 ExcelSheet.Application.Quit();	
  //  // ���񪫥��ܼơC						
 ExcelSheet = "";									
 // ExcelSheet.Delete;	
   }
//document.write(DATCld(2012,11,25));
 //document.write (SMM);
// document.write (TDD);�|���_�b�WResponse.Redirect 
 //document.write (SYY);
// document.write (SMM);
// document.write (TDD);


  }
  
  </script> 
 
<% 
  
 
 function tst()
   
   if ((&had) and (&h88))<>0 then
      'tst="hellow; ok"
      tst=Hex((&had) and (&h88))
      else 
      tst="hellow; no"
   end if
   End function   
  
  function LDy(yt,m) 
    
    hxy =("&h"&Mid(yt,1,5))
    'hxy =(yt&m)
    if ( hxy and (&h10000)) >m then
       LDy=Hex(hxy and (&h10000))&"hellow; ok"
      ''tst=Hex((&had) and (&h88))
      else 
       LDy=Hex(hxy and (&h10000))&"hellow; no"
   end if
  '' LDy=hxy
   ''return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );('0x'+sr)
  end function

 '' Response.Write "----------LYGG------"
 '' Response.Write tst()
 ''  Response.Write LDy("dacdf",6)
   %>
</html>

<%  
Function IsSelected( Which, Lesson ) 
   If Which = Lesson Then IsSelected = Selected 
End Function 
%>