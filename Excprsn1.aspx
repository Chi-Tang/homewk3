<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>
 

<%

  Dim  SSEX 
  Dim YYNUM, MMNUM ,DDNUM,HHNUM,LLYR 
  Dim NYNUM, NMNUM ,NDNUM,NHNUM 
  
  SSex = Request("Sex")
 YYNUM=Request("YNUM")
 MMNUM=Request("MNUM")
 DDNUM=Request("DNUM")
 HHNUM=Request("HNUM")
 ''LLYR=Request("LYR")
  'YKK=Request("YKK")
  'YKG=Request("YKG")
  'LYG=Request("LYG")
 Dim A(220)
 Dim B(220)
 Dim AW(220)
 Dim AKI(220)
 
 Dim ZLG()={"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
 Dim ZLN() = {"�g�T","����","�S�s","�妱","�G�s","�Z��","�}�x","�Z��","�G�s","�妱","�S�s","����"}
 Dim ZBN() = {"���P","�Ѭ�","�ѱ�","�ѦP","���","�Ѿ�","���P","�Ѭ�","�ѱ�","�ѦP","���","�Ѿ�"}
 
 Dim ZNAM() = {"�[","��","��","�S","�_","��","�I","��","�["}
 Dim ZNUM() = {"0,0,0","0,0,1","0,1,0","0,1,1","1,0,0","1,0,1","1,1,0","1,1,1","0,0,0"}
 Dim ZNAMM() = {"�a","�s","��","��","�p","��","�A","��","�a"}
 Dim ZK()={"��","�A","��","�B","��","�v","��","��","��","��"}
 Dim ZG()={"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
   ''Dim YKG=GRNDSN(SECTM, TNUM, DNUM, HNUM)
  '' Dim LYG=LRNDSN(LLYR, "2", "20")
  Dim YKN=KRNDSN(YYNUM, MMNUM, DDNUM, HHNUM)
  Dim YKG=GRNDSN(YYNUM, MMNUM, DDNUM, HHNUM)
    
 %>
 
<HTML> 
  <!-- <span style="writing-mode:tb-rl">�峹���e RowT</span>-->
 <style type="text/css">
  ''body  {
    ' writing-mode: tb-rl;direction: rtl;
    ' unicode-bidi: embed;background-color: blue;}
  td  {display: td;font-size: 11pt; writing-mode:tb-rl; 
        background-color: yellow; Border: non ;}
  tr  {display: tr;font-size: 11t; writing-mode:tb-rl; 
        background-color: yellow;Border: non  ;}
      
  table  {display: table;font-size: 11pt;writing-mode: tb-rl;
          border-collapse= collapse;Border: non ;}

 '' div {display: -ms-box;position:relative; top:40px+20px; width:130px+20px;
       Border: solid 1px ;writing-mode:tb-rl;
       background-color: red;column-count:4; -ms-grid-row: 4; }
  </style>
 
 <BODY bgcolor="#FFFFFF">
  <CENTER><H2>�� �L �� �� �R �L </H2>
 <%
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim rs3 As OleDbDataReader
      Dim SQL As String, Body As String
      Dim mad11, mad22 As String
      Dim k, n, j As integer
       k=0
      Dim MYL,MYB  
      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
        '''Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
     Dim Database = "Data Source=" & Server.MapPath("/HSU-Tanwen/ZUWE01.mdb")
      Dim Dbpass = "Jet OLEDB:Database Password=tang1206"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()
 
   'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
     SQL = "Select * From �R�c�� Where ����='" & MMNUM & "'"
       '' SQL = "Select * From �R�c��"
      Cmd = New OleDbCommand( SQL, Conn )
      rs3 = Cmd.ExecuteReader()
    ''While Rd.Read()
   If rs3 Is Nothing Then
   '''Response.Write ("GetExcelRecordset �I�s����!")
    Response.End
   End If 
  Dim z3(16)
   z3(0)="�R�c"
  Dim z3k =" "
  Dim wk=" "
   ''' Part II�G��X��ƪ��u���e�v
 While rs3.Read()
      '''If rs3.Item(0)= MMNUM Then ' ��ܦ��@ Email �s�b
       for k = 1 to rs3.Fieldcount -1
        '' mad11 = Rd.Item("��")
          if rs3.GetName(k) = HHNUM then
               
                 z3(k)=rs3.Item(k)
                 z3k = rs3.Item(k)
            
          End If 
        next 
   '' end if
  
  End While   
     rs3.close()
      ''Conn.Close()
   
  A(5)=z3(0)
  B(5)=z3k
  AW(5)=z3(0)&wk
  AKI(5)="K5"
     For n=0 to 11
        if z3k=ZLG(n) then
             MYL=ZLN(n)
        end if 
    Next   
  
     For j=0 to 11
        if Trim(YKG)=ZLG(j) then
           ''zgtk1=n
          MYB=ZBN(j)
        end if 
     Next
  
  '''Response.Write (MYB)
  '''Response.Write (MYL)
  '''Response.Write (A(5)&B(5))
   
 '''''''''''''''''''''''''''''''''''''''''''''''''
   
'''''''''''''''''''''''''''''''''''''''''''
  Dim rsb1 As OleDbDataReader 
  '' SQL = "Select * From �R�c��"  
    SQL = "Select * From ���c�� Where ����='" & MMNUM & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      rsb1 = Cmd.ExecuteReader()
    ''While Rd.Read()
   If rsb1 Is Nothing Then
    Response.Write ("GetExcelRecordset �I�s����!")
    Response.End
   End If 

  Dim zb1(16)
  zb1(0)="���c"
  Dim zb1k=" "
   ''Dim wk=" "
  Dim i 
   ''' Part II�G��X��ƪ��u���e�v
 While rsb1.Read()
   '' IF Trim(rsb1(0))= Trim(MMNUM)  Then
       
       For i=1 to rsb1.Fieldcount-1
        if rsb1.GetName(i)= HHNUM then
          'if rsb1(i).Name= "��" then
           zb1(i)=rsb1(i)
           zb1k =rsb1(i)
         End If 
       Next
   '' End If  
     
 End While
  rsb1.close()
     '' Conn.Close()

 A(3)=zb1(0)
 B(3)=zb1k
 AW(3)=zb1(0)&wk
 AKI(3)="K3"
 
'''Response.Write (A(3)&B(3))
  Dim rs5 As OleDbDataReader     
   SQL = "Select * From �l�c��"  
   '' SQL = "Select * From �l�c�� Where ����='" & MMNUM & "'"
      Cmd = New OleDbCommand( SQL, Conn )
      rs5 = Cmd.ExecuteReader()
    ''While Rd.Read()
   If rs5 Is Nothing Then
    Response.Write ("GetExcelRecordset �I�s����!")
    Response.End
   End If 
 Dim z5(16)
  z5(0)="�l�c��"
 Dim  z5k=" "
  k=5
  wk=" "

  '' Part II�G��X��ƪ��u���e�v
   While rs5.Read()	
          k=k+1
        z5(0)=rs5(0)
       For i=1 to rs5.FieldCount-1
         if rs5.GetName(i)= z3k then
           z5(i)=rs5(i)
           z5k =rs5(i)
          End If 
       Next
  
  A(0+k)=z5(0)
  B(0+k)=z5k
  AW(0+k)=z5(0)&wk
  AKI(0+k)="K"&k
 End While
  
  rs5.close()
     '' Conn.Close()

  '''Response.Write (A(7)&B(7))
 '''''''''''''''''''''''''''''''''''''''''''''
  Dim rs2 As OleDbDataReader     
   SQL = "Select * From ���槽 Where ����='" & YKN & "'"
   Cmd = New OleDbCommand( SQL, Conn )
      rs2 = Cmd.ExecuteReader()
   
If rs2 Is Nothing Then
    Response.Write ("GetExcelRecordset �I�s����!")
    Response.End
End If 
 Dim z2(16)
  z2(0)="(��)�R�c"
 Dim z2k=" "
  wk=" "
  '' Part II�G��X��ƪ��u���e�v
  While rs2.Read()	

   '' IF Trim(rs2(0))= Trim(YKN)  Then 
       For i=1 to rs2.FieldCount-1
          if rs2.GetName(i)= z3k then
             z2(0)=rs2(i)
             z2k =rs2(i)
           End If 
       Next
    
   '' End If  
    
  End While
   
  A(2)=z2(0)
  B(2)=z2k
  AW(2)=z2(0)&wk
  AKI(2)= "K2"
   rs2.close()
     '' Conn.Close()

  '''Response.Write (A(2)&B(2)&"<br>")


''''''''''''''''''''''''''''''''
 Dim rs1 As OleDbDataReader     
   ''SQL = "Select * From ���L�� Where ����='" & DDNUM & "'"
    SQL = "Select * From ���L��"
   Cmd = New OleDbCommand( SQL, Conn )
      rs1 = Cmd.ExecuteReader()
  If rs1 Is Nothing Then
    Response.Write ("GetExcelRecordset �I�s����!")
    Response.End
  End If 
  
 Dim z1(16)
  z1(0)="���L�P"
 Dim z1k=" "
 Dim z11k=" "
 Dim ZB() = {"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
  Dim ZF() = {"��","�f","�G","��","�l","��","��","��","��","��","��","�w"}
 
  '' Part II�G��X��ƪ��u���e�v
  While rs1.Read()	

  IF rs1.Item(0) = DDNUM  Then 
       
       For i=1 to rs1.FieldCount-1
         if rs1.GetName(i)= z2k then
           z1(i)=rs1.Item(i)
           z1k =rs1.Item(i)
           For j=0 to 11
            ''if ZB(j)= rs1.Item(i) then
             if ZB(j)= z1k then
               z11k = ZF(j)
             End If 
           Next
       End If 
       Next
    
   End If  
    
  End While
 rs1.close()
     '' Conn.Close()
  ' A(1)=z1(0)
 ' B(1)=z1k
 ' AW(1)=z1(0)&wk
  'AKI(1)= "K1"
 ''''Response.Write (A(1))  
  '''Response.Write (z1(0))   
  '''Response.Write (z1k)
  '''Response.Write (z11k)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   '' Dim k, i, j As integer
    Dim rs As OleDbDataReader
      SQL = "Select * From �ҬP�L"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"

      Cmd= New OleDbCommand( SQL, Conn )
      rs = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rs.Read()= False Then
    If rs Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
      ''''Response.Write(rs.GetName(0))
      ''''Response.Write (rs(1))
    Response.End
   End If 
  Dim z(16)
   z(0)="(���L)�ѩ��c"
  Dim zzk=" "
  k=20
 
  '' Part II�G��X��ƪ�k=k+1,add���e�v
   While rs.Read()
      k=k+1
        z(0)=rs(0)
    For j=1 to rs.FieldCount-1
       if rs.GetName(j)= z1k then
        ''if rs.GetName(j)= "��" then
           z(j)=rs.Item(j)
           zzk =rs.Item(j)
         End If 
       Next
  
  '' wkw=WRNDSN(z(0), zzk)
    A(0+k)=z(0)
  B(0+k)=zzk
  ''AW(0+k)=z(0)&wkw
   
  End While 
    ''rs.close()
    '' Conn.Close()
   '''Response.Write (A(23))    
   '''Response.Write (B(23)&"<br>")
  For i= 21 to 27
   Dim zof= A(0+i)
   Dim zzkf=B(0+i)
    ''wkw=WRNDSN(z(0), zzk)
   Dim wkw=WRNDSN(zof,zzkf)
    AW(0+i)=A(0+i)&wkw
 '''Response.Write (Aw(i))
  Next
   rs.close()
     '' Conn.Close()

  ''Response.Write (Aw(23))

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
   Dim rsf As OleDbDataReader
      SQL = "Select * From �ҬP�L2"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rsf = Cmd.ExecuteReader()
      ''While Rd.Read()
      ''If rd.Read()= False Then
    If rsf Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
      Response.End
   End If 
  
  Dim zff(16)
 ' zff(0)="(���L)�ѩ��c"
 Dim zffk=" "
  k=30
  ' Part II�G��X��ƪ��u���e�v
  While rsf.Read()	
     k=k+1
   
     zff(0)=rsf(0)
       For i=1 to rsf.FieldCount-1
         if rsf.GetName(i)= z11k then
           zff(i)=rsf.Item(i)
           zffk =rsf.Item(i)
         End If 
       Next
   
  '' wkw=WRNDSN(zff(0), zffk)
  
  A(0+k)=zff(0)
  B(0+k)=zffk
  ''AW(0+k)=zff(0)&wkw
 
 End While
  
'''Response.Write (A(33))    
   '''Response.Write (B(33)&"<br>")
  For i= 31 to 37
   Dim zfko= A(0+i)
   Dim zfkk=B(0+i)
    ''wkw=WRNDSN(z(0), zzk)
   Dim wkw=WRNDSN(zfko,zfkk)
    AW(0+i)=A(0+i)&wkw
 '''Response.Write (Aw(i))
  Next
   rsf.close()
     '' Conn.Close()

  ''Response.Write (Aw(33))
 
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim SexYK ,SexYKT
 Select Case Trim(YKN)
     Case "��","��","��","��","��"
         SexYk="��"&SSex
           Select Case Trim(SSex)
             Case "�k"
                 SexYKT="S"
             Case "�k"
                 SexYKT="R"
            ' Case Else
              ' SexYKT="N""
             End  Select   
         
     Case "�A","�B","�v","��","��"
         SexYk="��"&SSex
            Select Case Trim(SSex)
             Case "�k"
                 SexYKT="R"
             Case "�k"
                 SexYKT="S"
            ' Case Else
              ' SexYKT="N""
             End  Select 
     Case Else
         SexYk="��"&SSex
    End  Select   
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rsft As OleDbDataReader
      SQL = "Select * From �ҬP�L3"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rsft = Cmd.ExecuteReader()
      ''While Rd.Read()
      ''If rd.Read()= False Then
    If rsft Is Nothing Then
       Response.Write ("GetExcelRecordset �I�s����!")
       Response.End
    End If 
   Dim zft(16)
   '' zft(0)="�S�s�P"
  Dim zftk=" "
  Dim subtk
   k=40
  '' Part II�G��X��ƪ��u���e�v
   While rsft.Read()
     k=k+1
    zft(0)=rsft(0)
       For i=1 to rsft.FieldCount-1
        if rsft.GetName(i)= Trim(YKN) then
           zft(i)=rsft.Item(i)
           zftk =rsft.Item(i)
            IF Trim(rsft(0))= "�S�s"  Then 
             subtk=zftk 
              
             '' DOCTOR SexYKT, subtk
           End If
       End If  
     Next
  Dim wkw=WRNDSN(zft(0), zftk)
  A(0+k)=zft(0)
  B(0+k)=zftk
   ''AW(0+k)=zft(0)& wkw
  
 End While
 
'''Response.Write (A(3))    
   '''Response.Write (B(43)&"<br>")
  For i= 41 to 47
   Dim zfto= A(0+i)
   Dim zftko= B(0+i)
    ''wkw=WRNDSN(z(0), zzk)
   Dim wkw=WRNDSN(zfto,zftko)
    AW(0+i)=A(0+i)&wkw
 '''Response.Write (Aw(i))
  Next
   rsft.close()
     '' Conn.Close()

  ''Response.Write (Aw(43))

 
 '' DOCTOR SexYKT, subtk
'' SUB DOCTOR(SexYKT, subtk)
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim rspt As OleDbDataReader
      SQL = "Select * From �դh��"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rspt = Cmd.ExecuteReader()
      ''While Rd.Read()
      ''If rd.Read()= False Then
    If rspt Is Nothing Then
       Response.Write ("GetExcelRecordset �I�s����!")
       Response.End
    End If 
  Dim zpt(24)
   '' zpt(0)="�S�s�P"
   wk=" "
   Dim zptk
   Dim Sxtftk=SexYKT & subtk
   k=75
 '' Part II�G��X��ƪ��u���e�v
  While rspt.Read() 
     k=k+1
    zpt(0)=rspt(0)
       For i=1 to rspt.FieldCount-1
        if rspt.GetName(i)= Trim(Sxtftk) then
         'if rspt(i).Name= "��" then
           'zpt(0)="�S�s�P"
           zpt(i)=rspt.Item(i)
           zptk =rspt.Item(i)
         End If 
       Next
   
  A(0+k)=zpt(0)
  B(0+k)=zptk
  AW(0+k)=zpt(0)&wk
  'AW(0+k)=zpt(0)&wk
End While
 
   For i= 76 to 87
   Dim zpto= A(0+i)
   Dim zptko= B(0+i)
    ''wkw=WRNDSN(z(0), zzk)
   Dim wkw=WRNDSN(zpto,zptko)
    AW(0+i)=A(0+i)&wkw
 '''Response.Write (Aw(i))
  Next
   rspt.close()
     '' Conn.Close()
 '''Response.Write (A(76))    
   '''Response.Write (B(76)&"<br>") 
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
   Dim rs6 As OleDbDataReader
      SQL = "Select * From ���P��"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rs6 = Cmd.ExecuteReader()
      ''While Rd.Read()
      ''If rd.Read()= False Then
    If rs6 Is Nothing Then
       Response.Write ("GetExcelRecordset �I�s����!")
       Response.End
    End If 
  Dim z6(16)
  z6(0)="���P"
  Dim z6k=" "
  k=51
  '''Part II�G��X��ƪ��u���e�v
 While rs6.Read()
     z6(0)="���P"
   IF Trim(rs6(0))= Trim(YKG)  Then 
     For i=1 to rs6.FieldCount-1
        if rs6.GetName(i)= Trim(HHNUM) then
           z6(i)=rs6.Item(i)
           z6k =rs6.Item(i)
         
        End If 
       Next
    End If 
  A(0+k)=z6(0)
  B(0+k)=z6k
 '' AW(0+k)=z(0)&wkw
  End While 
   rs6.close()
     '' Conn.Close()
  '' '''Response.Write (A(51)&"<br>")  
 
 ''Dim wkwg=WRNDSN(z6(0), z6k)
  '' For i= 21 to 26
   Dim zfof= A(0+k)
   Dim zfok=B(0+k)
    ''WKW=WRNDSN(zzo, zftk)
   '''Dim wkwf=WRNDSN(z6(0), z6k)
  Dim wkwf=WRNDSN(zfof,zfok)
    AW(0+k)=A(0+k)&wkwf
 '''Response.Write (Aw(k))
 '' Next
    
  
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   
   Dim rs7 As OleDbDataReader
     SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
       '' SQL = "Select * From �a�P��"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rs7 = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rs7.Read()= True Then
    If rs7 Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
      '' rs7.close()
      '' Conn.Close()
    Response.End
   End If 
  Dim z7(16)
   z7(0)="�a�P"
  Dim z7k=" "
   k=52
  '' Part II�G��X��ƪ��u���e�v
  While rs7.Read() 
      '' If rs7.IsDbNull(0) = True Then
   z7(0)="�a�P"
     '' IF Trim(rs7.GetName(0)) = Trim(YKG)  Then 
       For i=1 to rs7.FieldCount-1
        if rs7.GetName(i) = Trim(HHNUM) then
           z7(i)=rs7.Item(i)
           z7k =rs7.Item(i)
         End If 
       Next
   '' End If  
  ''' Dim wkw=WRNDSN(z(0), zzk)
  A(0+k)=z7(0)
  B(0+k)=z7k
 '' AW(0+k)=z(0)&wkw
 End While 
   rs7.close()
     '' Conn.Close()
  ''  Response.Write (A(52)&"<br>")  
 ''Dim wkwg=WRNDSN(z7(0), z7k)
  '' For i= 21 to 26
   Dim zofg= A(0+k)
   Dim zzkg=B(0+k)
    ''wkw=WRNDSN(z(0), zzkf)
  Dim wkwg=WRNDSN(zofg,zzkg)
    AW(0+k)=A(0+k)&wkwg
 '''Response.Write (Aw(k))
 '' Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    Dim rsf7 As OleDbDataReader
      SQL = "Select * From �ҬP�L4"
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rsf7 = Cmd.ExecuteReader()
      ''While Rd.Read()
      ''If rd.Read()= False Then
    If rsf7 Is Nothing Then
       Response.Write ("GetExcelRecordset �I�s����!")
       Response.End
    End If 
 Dim zf7(16)
 ' zf7(0)="���(��)�P"
 Dim zf7k=" "
  k=53
  ' Part II�G��X��ƪ��u���e�v
  While rsf7.Read()
     k=k+1
     zf7(0)=rsf7(0)
       For i=1 to rsf7.FieldCount-1
        if rsf7.GetName(i)= Trim(HHNUM) then
         
           zf7(i)=rsf7.Item(i)
           zf7k =rsf7.Item(i)
         End If 
       Next
   
  '' wkw=WRNDSN(zf7(0), zf7k)
  A(0+k)=zf7(0)
  B(0+k)=zf7k
  ''AW(0+k)=zf7(0)&wkw
 End While
   rsf7.close()
     '' Conn.Close()
  '''Response.Write (A(54)&"<br>")  
 
   For i= 54 to 59
   Dim zofh= A(0+i)
   Dim zzkh=B(0+i)
    ''wkw=WRNDSN(z(0), zzkf)
  Dim wkwh=WRNDSN(zofh,zzkh)
    AW(0+i)=A(0+i)&wkwh
 '''Response.Write (Aw(i))
  Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim rsgt As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �~��P��" 
       '' SQL = "Select * From �ҬP�L Where ����='" & "���L" & "'"
      Cmd= New OleDbCommand( SQL, Conn )
      rsgt = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rsgt.Read()= True Then
    If rsgt Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
  Dim zgt(16)
   '' zgt(0)="�Ѥ~��"
  Dim zgtk=" "
   wk=" "
  k=90
 '' Part II�G��X��ƪ��u���e�v
 While rsgt.Read()= True
    If rsgt.IsDbNull(0) = True Then
      ''Response.Write ("����� �I�s����!")
       Exit While
    End If 
  
   k=k+1
    zgt(0)=rsgt.Item(0)
       For i=1 to rsgt.FieldCount-1
       if rsgt.GetName(i) = Trim(YKG) then
           zgt(i)=rsgt.Item(i)
           zgtk =rsgt.Item(i)
         
         IF Trim(rsgt.Item(0))= "�Ѥ~"  Then 
            For n=5 to 17
               if rsgt(i)=A(n) then
                  zgtk=B(n)
               end if                
             Next
          End If
           IF Trim(rsgt.Item(0))= "�ѹ�"  Then 
                ' zgtk= GGRNDSN(B(3), YKG)
                Dim  BN3=Trim(B(3))
                Dim  YKGS=Trim(YKG)
                zgtk= GGRNDSN(BN3, YKGS)
             '' GGRNDSN BN3, YKGS

            End IF
       End If  
     Next
  
 A(0+k)=zgt(0)
  B(0+k)=zgtk
  AW(0+k)=zgt(0)& wk
  'AW(0+k)=zgt(0)&wk
'''Response.Write (A(k))
'''Response.Write (B(k))
 
 End While  
 rsgt.close()
     '' Conn.Close()
   '''Response.Write (A(103)&"<br>")  
 ''Dim wkwg=WRNDSN(z7(0), z7k)
  '' Dim zgtk= GGRNDSN(B(3), YKG)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rsmt As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From ��t�P��" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsmt = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rsmt.Read()= True Then
    If rsmt Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
   
 Dim zmt(16)
  '' zmt(0)="���k�]"
 Dim zmtk=" "
 Dim TDDNUM1,TDDNUM2,TDDNUM3,TDDNUM4
  wk=" "
  k=60
 '' Part II�G��X��ƪ��u���e; �P�_�O�_�L�F�̫�@���v
  
 While rsmt.Read()= True
    '' if rsmt.Read() = True then
   If rsmt.IsDbNull(0) = True Then
      ''Response.Write ("����� �I�s����!")
       Exit While
    End If 
 
     k=k+1
     zmt(0)=rsmt.Item(0)
    For i=1 to rsmt.FieldCount-1
        if Trim(rsmt.GetName(i))=Trim("M"&MMNUM) then
         'if rsmt(i).Name= "��" then
           'zmt(0)="���k�]"
           zmt(i)=rsmt(i)
           zmtk =rsmt(i)
           
      Select Case Trim(zmt(0))
        Case "�T�x"
          TDDNUM1=DDNUM-1
          zmtk=TGRNDSN(B(61), TDDNUM1)
           ' DOCTOR SexYKT, subtk
         Case "�K�y" 
           TDDNUM2=-(DDNUM-1)
           zmtk=TGRNDSN(B(62), TDDNUM2)
            ' GGRNDSN BN5, YKGS
         Case "����"
          TDDNUM3=DDNUM-2
          zmtk=TGRNDSN(B(54), TDDNUM3)
           ' DOCTOR SexYKT, subtk
         Case "�ѶQ" 
           TDDNUM4=DDNUM-2
           zmtk=TGRNDSN(B(55), TDDNUM4)
            ' GGRNDSN BN5, YKGS
       End  Select      
      
      End If  
    Next
 
  A(0+k)=zmt(0)
  B(0+k)=zmtk
  AW(0+k)=zmt(0)& wk
  'AW(0+k)=zmt(0)&wk
'''Response.Write (A(k))
'''Response.Write (B(k))

 End While
  rsmt.close()
     '' Conn.Close()
   '' Response.Write (A(63)&"<br>") 
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim rssk As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From ���Ū�" 
      Cmd= New OleDbCommand( SQL, Conn )
      rssk = Cmd.ExecuteReader()
   If rssk Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
 Dim zsk(16)
   '' zsk(0)="���Ŧ~��"
 Dim zskk=" "
  wk=" "
  k=105
 '' Part II�G��X��ƪ��u���e; �P�_�O�_�L�F�̫�@���v
 While rssk.Read()= True
   If rssk.IsDbNull(0) = True Then
      ''Response.Write ("����� �I�s����!")
       Exit While
    End If 
  if  Trim(rssk.Item(0))= Trim(YKG)  then
        '' zsk(0)=rssk(0)
       For i=1 to rssk.FieldCount-1
        if Trim(rssk.GetName(i))= Trim(YKN)  then
           zsk(i)=rssk.Item(i)
           zskk =rssk.Item(i)
         End If 
       Next
  
    ''A(0+k)=zsk(0)
   A(0+k)="����"
   B(0+k)=zskk
   ''AW(0+k)=zsk(0)&wk
   AW(0+k)="����"&wk
  End if  
 
 ''Response.Write (A(k))
 ''Response.Write (B(k))

 End While
  rssk.close()
     '' Conn.Close()
   '''Response.Write (A(105)) 
   '''Response.Write (B(105)&"<br>") 

 A(106)="�Ѷ�"
   B(106)=B(12)
   AW(106)="�Ѷ�"&wk
   'A(12)="���Юc"
 A(107)="�Ѩ�"
   B(107)=B(10)
   AW(107)="�Ѩ�"&wk
   ''A(10)="�e�̮c"

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rsk As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �|�ƬP��" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsk = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rsmt.Read()= True Then
    If rsk Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
 Dim zfk(16)
   '' zfk(0)="�|�ƬP�c"
 Dim zfk4=" "
  wk=" "
  k=141
 '' Part II�G��X��ƪ��u���e�v
  While rsk.Read()= True
    '' if rsk.Read() = True then
   If rsk.IsDbNull(0) = True Then
      ''Response.Write ("����� �I�s����!")
       Exit While
    End If 
     k=k+1
    zfk(0)=rsk.Item(0)
       For i=1 to rsk.FieldCount-1
        if rsk.GetName(i)= Trim(YKN)  then
           zfk(i)=rsk.Item(i)
           zfk4 =rsk.Item(i)
         
  '''''IF Trim(rsgt.Item(0))= "�Ѥ~"  Then 
         '   For n=21 to 62
         '     On Error Resume Next 
          '    if zfk(i)=A(n) then
              
          '       If Err.Number = 0 Then 
          '         zfk4 =B(n)
          '        '  Else
          '        ' Response.Write (Err.Number)
           '        ' Response.Write ("<br>")
          '       End If
          '  '' '' zfk4 =B(n)
          '     end if                
          '  Next
  '''''''End If
      End If 
       Next
   A(0+k)=zfk(0)
   B(0+k)=zfk4
   AW(0+k)=zfk(0)&wk
 
 '''Response.Write (A(k))
 '''Response.Write (B(k))
 End While
  rsk.close()
     '' Conn.Close()
  ''  Response.Write (A(105)&"<br>") 
  ''  Response.Write (B(105)&"<br>") 

 '''''''''''''''''''''''''''''''''''''''''
    Dim rsg As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �j����" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsg = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rsmt.Read()= True Then
    If rsg Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
  Dim zfl(24)
  '' zfl(0)="�j����c"
  Dim zflk=" "
  wk=" "
  k=145
 '' Part II�G��X��ƪ��u���e�v
  While rsg.Read()= True
     If rsg.IsDbNull(0) = True Then
      ''Response.Write ("����� �I�s����!")
       Exit While
    End If 
   k=k+1
    zfl(0)=rsg.Item(0)
       For i=1 to rsg.FieldCount-1
        if  rsg.GetName(i)= Trim(SexYK+z2k)  then
           zfl(i)=rsg.Item(i)
           zflk =rsg.Item(i)
        '''  '''  IF Trim(rsgt.Item(0))= "�Ѥ~"  Then 
           '  For n=5 to 17
           '    if zfl(0)=A(n) then
            '      ''zflk =B(n)
            '      B(0+k)=B(n)

           '    end if                
           '  Next
   '''       End If


         End If 
       Next
  
   B(0+k)=zfl(0)
   A(0+k)=zflk
   AW(0+k)=zflk
  ''AW(0+k)=zflk & wk

 '''Response.Write (A(k))
 '''Response.Write (B(k))
 End While
  rsg.close()
    ''Conn.Close()
  ''  Response.Write (A(105)&"<br>") 
  ''  Response.Write (B(105)&"<br>") 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rsgs As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �p����D" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsgs = Cmd.ExecuteReader()
     ''While Rd.Read()
     ''If rsmt.Read()= True Then
    If rsgs Is Nothing Then
     Response.Write ("GetExcelRecordset �I�s����!")
     Response.End
   End If 
  Dim zfms(24)
   '' zfms(0)="�p����c"
  Dim zfmsk=" "
  wk=" "
  k=160
 '' Part II�G��X��ƪ��u���e�v
  While rsgs.Read()= True
     If rsgs.IsDbNull(0) = True Then
      Response.Write ("����� �I�s����!")
       rsgs.close()
    Conn.Close()
   Exit While
    End If 
  k=k+1
   zfms(0)=rsgs.Item(0)
       For i=1 to rsgs.FieldCount-1
        if  rsgs.GetName(i)= Trim(SSex+YKG)  then
           zfms(i)=rsgs.Item(i)
           zfmsk =rsgs.Item(i)
    
     End If 
       Next
    
   A(0+k)= zfms(0)
   B(0+k)=zfmsk
   AW(0+k)=zfms(0)&wk
  '''Response.Write (A(k))
 '''Response.Write (B(k))
 End While
  rsgs.close()
     ''Conn.Close()
   '' Response.Write (A(105)&"<br>") 
   '' Response.Write (B(105)&"<br>") 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim RKS(15,6)
 
 Dim R() = {" ","[�l�c]","[���c]","[�G�c]","[�f�c]","[���c]","[�w�c]","[�Ȯc]","[���c]","[�Ӯc]","[���c]","[���c]","[��c]","[�� �� �c]"}

 Dim  RE() = {" "," "," "," "," "," "," "," "," "," "," "," "," "," "}

 Dim KFE() = {" "," "," "," "," "," "," "," "," "," "," "," "," "," "}

 Dim  KF() = {" ","[�l�c]","[���c]","[�G�c]","[�f�c]","[���c]","[�w�c]","[�Ȯc]","[���c]","[�Ӯc]","[���c]","[���c]","[��c]","[�� �� �c]"}

 Dim  KS() = {"�R�c","�S��","�ҩd","�l�k","�]��","�e��","�E��","����","�x�S","�Цv","�ּw","����","���L","�Ѿ�","�Ӷ�","�Z��","�ѦP","�G�s","�ѩ�","�ӳ�","�g�T","����","�Ѭ�","�ѱ�","�C��","�}�x","�ƸS","���v","�Ƭ�","�Ƨ�"," "}
 Dim  KI() = {"K0","K1","K2","K3","K4","K5","K6","K7","K8","K9","K10","K11","S0","S1","S2","S3","S4","S5","T0","T1","T2","T3","T4","T5","T6","T7"," "," "," "," "," "}

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''�����c�W 
 Dim BBG ,AAJ ,AAJm, AAN,BBN
 Dim m 
   On Error Resume Next 
   For j=142 to 162
      BBG=Trim(B(j))
      AAJ=Trim(A(j))
   
     If B(J) Is Nothing Then
        ''Response.Write ("�����B(J), �I�s����!")
       Exit For
     End If 
  For n= 1 to 200 
       AAN=Trim(A(n))
       BBN=Trim(B(n))  
   If BBG=AAN then
       If Err.Number = 0 Then 
             BBG=AAN 
           '     Else
            ' Response.Write (Err.Number)
        End If
   AAJm = Mid(Trim(AAJ),1,1)
    Select Case AAJm
      Case "��","��"
       AW(j)=Trim(BBG)&Trim(AW(j))
      Case Else
       AW(j)=Trim(AW(j))
    End Select   
     ''B(j)=B(n)
    B(j)=BBN
        If Err.Number = 0 Then 
             B(j)=BBN
          ' Else
          '     Response.Write (Err.Number)
         End If
     End if 
   Next
  '''Response.Write (Aw(j)) 
  '''Response.Write (B(j)&"<br>") 
 Next  

 '''''IF Trim(rsgt.Item(0))= "�|�Ƥj���Ѥ~"  Then 
       '     For n=21 to 62
       '       On Error Resume Next 
        '      if zfk(i)=A(n) then
              
        '         If Err.Number = 0 Then 
         '          zfk4 =B(n)
         '           Else
         '          Response.Write (Err.Number)
         '           Response.Write ("<br>")
          '       End If
          '    end if                
          '  Next
  '''''''End If

   
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim rsy4 As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �c�z��" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsy4 = Cmd.ExecuteReader()
    
     ''If rsy4.Read()= True Then
    If rsy4 Is Nothing Then
      Response.Write ("GetExcelRecordset �I�s����!")
      Response.End
    End If 
   Dim zt(12)
    'zt(0)="�c�z"
   Dim ztk=" "

   '' Part II�G��X��ƪ��u���e�v
    On Error Resume Next 
   While rsy4.Read()
     If rsy4.Item(0) = Trim(YKN) Then
     
       For i = 1 To rsy4.FieldCount - 1
         zt(i) = rsy4.Item(i) + rsy4.GetName(i)
         R(i) = "[" + zt(i) + "�c]"
   '''Response.Write (R(i))
       Next
     End If
  End While
  rsy4.close()
     '' Conn.Close()
   '' Response.Write (A(105)&"<br>") 
   '' Response.Write (B(105)&"<br>") 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsy5 As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �y�~�ѬP��" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsy5 = Cmd.ExecuteReader()
    
     ''If rsy5.Read()= True Then
    If rsy5 Is Nothing Then
      Response.Write ("GetExcelRecordset �I�s����!")
      Response.End
    End If 
  
  Dim zy5(16)
   '' zy5(0)="�y�~�R�c�ѬP"
  Dim zy5k=" "
   k=175
   wk=" "
   '' Part II�G��X��ƪ��u���e�v
    On Error Resume Next 
   While rsy5.Read()
      k=k+1
      zy5(0)=rsy5.Item(0)
       For i=1 to rsy5.FieldCount-1
         if rsy5.GetName(i)= Trim(YKG) then
           ''if rsy5.GetName(i)= Trim(LYG) then
           zy5(i)=rsy5.Item(i)
           zy5k =rsy5.Item(i)
          
        End If 
       Next
  A(0+k)=zy5(0)
  B(0+k)=zy5k
  AW(0+k)=zy5(0)&wk
 
 '''Response.Write (A(k)) 
 '''Response.Write (B(k)) 

 End While
  rsy5.close()
     '' Conn.Close()
   '' Response.Write (A(105)&"<br>") 
   '' Response.Write (B(105)&"<br>") 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim rsy6 As OleDbDataReader
     '' SQL = "Select * From �a�P��  Where ����='" & Trim(YKG)  & "'"
      SQL = "Select * From �y�~��g��" 
      Cmd= New OleDbCommand( SQL, Conn )
      rsy6 = Cmd.ExecuteReader()
    
     ''If rsy6.Read()= True Then
    If rsy6 Is Nothing Then
      Response.Write ("GetExcelRecordset �I�s����!")
      Response.End
    End If 
  
  Dim zy6(16)
   '' zy6(0)="�y�~��g�c"
  Dim zy6k=" "
  Dim zy66k=" "
   k= 212
   wk=" "
   '' Part II�G��X��ƪ��u���e�v
    On Error Resume Next 
   While rsy6.Read()
    If TRIM(rsy6.Item(0))=Trim(MMNUM) Then 
         ''zy6(0)=rsy6(0)
       For i=1 to rsy6.FieldCount-1
         ''if rsy6.GetName(i)= Trim(LYG) then
          if rsy6.GetName(i)= Trim(YKG) then
           zy6(i)=rsy6.Item(i)
           zy66k =rsy6.Item(i)
           zy6k= LGRNDSN(zy66k, HHNUM)
    
        End If 
       Next
    End If  
  
  ''A(0+k)=zy6(0)
  A(212)="����g"
  B(212)=zy6k
   '' AW(0+k)="����g"zy6(0)&wk
  AW(0+k)="����g"&wk
  
 '' Response.Write (A(k)) 
  ''Response.Write (B(k)) 
 End While
  rsy5.close()
      Conn.Close()
   '''Response.Write (A(212)) 
   '''Response.Write (B(212)&"<br>") 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' A(215)="*�{���ۧ@��:"
   ' B(215)="�� �� �c"
   'AW(215)="*�{���ۧ@��:"&wk
 '' AW(0+k)="����g"zy6(0)&wk
  Dim BBK
  For j=0 to 220
      BBK=Trim(B(j))
   Select Case BBK
      Case "�l"
        R(1)=Trim(R(1))+","& Trim(AW(j))
        RE(1)=Trim(RE(1))+Trim(AW(j))
         KFE(1)=KFE(1)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(1,1)=Trim(RKS(1,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(1,2)=Trim(RKS(1,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(1,3)=Trim(RKS(1,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(1,4)=Trim(RKS(1,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(1,5)=Trim(RKS(1,1))+Trim(A(j))
           Case Else
              RKS(1,6)=Trim(RKS(1,1))+Trim(A(j))
         End  Select      

    

       Case "��"
        R(2)=R(2)+","&AW(j)  
        RE(2)=Trim(RE(2))+Trim(AW(j))
         KFE(2)=KFE(2)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(2,1)=Trim(RKS(2,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(2,2)=Trim(RKS(2,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(2,3)=Trim(RKS(2,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(2,4)=Trim(RKS(2,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(2,5)=Trim(RKS(2,1))+Trim(A(j))
 
           Case Else
              RKS(2,6)=Trim(RKS(2,1))+Trim(A(j))
         End  Select      


       Case "�G"
        R(3)=R(3)+","&AW(j) 
         RE(3)=Trim(RE(3))+Trim(AW(j)) 
          KFE(3)=KFE(3)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(3,1)=Trim(RKS(3,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(3,2)=Trim(RKS(3,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(3,3)=Trim(RKS(3,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(3,4)=Trim(RKS(3,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(3,5)=Trim(RKS(3,1))+Trim(A(j))
           Case Else
              RKS(3,6)=Trim(RKS(3,1))+Trim(A(j))
         End  Select      

             
      Case "�f"
        R(4)=R(4)+","&AW(j)
        RE(4)=Trim(RE(4))+Trim(AW(j)) 
         KFE(4)=KFE(4)+AKI(j) 
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(4,1)=Trim(RKS(4,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(4,2)=Trim(RKS(4,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(4,3)=Trim(RKS(4,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(4,4)=Trim(RKS(4,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(4,5)=Trim(RKS(4,1))+Trim(A(j))
 
           Case Else
              RKS(4,6)=Trim(RKS(4,1))+Trim(A(j))
         End  Select      


      Case "��"
        R(5)=R(5)+","&AW(j)
        RE(5)=Trim(RE(5))+Trim(AW(j)) 
         KFE(5)=KFE(5)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(5,1)=Trim(RKS(5,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(5,2)=Trim(RKS(5,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(5,3)=Trim(RKS(5,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(5,4)=Trim(RKS(5,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(5,5)=Trim(RKS(5,1))+Trim(A(j))
           Case Else
              RKS(5,6)=Trim(RKS(5,1))+Trim(A(j))
         End  Select      


     Case "�w"
        R(6)=R(6)+","&AW(j) 
        RE(6)=Trim(RE(6))+Trim(AW(j))
         KFE(6)=KFE(6)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(6,1)=Trim(RKS(6,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(6,2)=Trim(RKS(6,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(6,3)=Trim(RKS(6,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(6,4)=Trim(RKS(6,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(6,5)=Trim(RKS(6,1))+Trim(A(j))
            Case Else
              RKS(6,6)=Trim(RKS(6,1))+Trim(A(j))
         End  Select      

    
      Case "��"
        R(7)=R(7)+","&AW(j) 
        RE(7)=Trim(RE(7))+Trim(AW(j))
         KFE(7)=KFE(7)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(7,1)=Trim(RKS(7,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(7,2)=Trim(RKS(7,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(7,3)=Trim(RKS(7,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(7,4)=Trim(RKS(7,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(7,5)=Trim(RKS(7,1))+Trim(A(j))
           Case Else
              RKS(7,6)=Trim(RKS(7,1))+Trim(A(j))
          End  Select      
       

       Case "��"
        R(8)=R(8)+","&AW(j)
        RE(8)=Trim(RE(8))+Trim(AW(j))
         KFE(8)=KFE(8)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(8,1)=Trim(RKS(8,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(8,2)=Trim(RKS(8,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(8,3)=Trim(RKS(8,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(8,4)=Trim(RKS(8,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(8,5)=Trim(RKS(8,1))+Trim(A(j))
           Case Else
              RKS(8,6)=Trim(RKS(8,1))+Trim(A(j))
         End  Select      

      Case "��"
        R(9)=R(9)+","&AW(j)
        RE(9)=Trim(RE(9))+Trim(AW(j)) 
          KFE(9)=KFE(9)+AKI(j)
            Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(9,1)=Trim(RKS(9,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(9,2)=Trim(RKS(9,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(9,3)=Trim(RKS(9,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(9,4)=Trim(RKS(9,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(9,5)=Trim(RKS(9,1))+Trim(A(j))
           Case Else
              RKS(9,6)=Trim(RKS(9,1))+Trim(A(j))
         End  Select      
           

       Case "��"
        R(10)=R(10)+","&AW(j) 
        RE(10)=Trim(RE(10))+Trim(AW(j)) 
         KFE(10)=KFE(10)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(10,1)=Trim(RKS(10,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(10,2)=Trim(RKS(10,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(10,3)=Trim(RKS(10,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(10,4)=Trim(RKS(10,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(10,5)=Trim(RKS(10,1))+Trim(A(j))
           Case Else
              RKS(10,6)=Trim(RKS(10,1))+Trim(A(j))
          End  Select      


       Case "��"
        R(11)=R(11)+","&AW(j)
        RE(11)=Trim(RE(11))+Trim(A(j))
         KFE(11)=KFE(11)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(11,1)=Trim(RKS(11,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(11,2)=Trim(RKS(11,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(11,3)=Trim(RKS(11,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(11,4)=Trim(RKS(11,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(11,5)=Trim(RKS(11,1))+Trim(A(j))
           Case Else
              RKS(11,6)=Trim(RKS(11,1))+Trim(A(j))
         End  Select      


      Case "��"
       R(12)=R(12)+","&AW(j)
       RE(12)=Trim(RE(12))+Trim(AW(j)) 
        KFE(12)=KFE(12)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(12,1)=Trim(RKS(12,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(12,2)=Trim(RKS(12,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(12,3)=Trim(RKS(12,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(12,4)=Trim(RKS(12,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(12,5)=Trim(RKS(12,1))+Trim(A(j))
           Case Else
            RKS(12,6)=Trim(RKS(12,1))+Trim(A(j))
          End  Select      

       Case Else
         ''R(13)= Trim(YYNUM)&Trim(MMNUM)&Trim(DDNUM)&Trim(HHNUM)&",�R�D:"&Trim(MYL)&",���D:"&Trim(MYB)
        R(13)=R(13)&AW(j)
        RE(13)=Trim(RE(13))+Trim(A(j)) 
         KFE(13)=KFE(13)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(13,1)=Trim(RKS(13,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(13,2)=Trim(RKS(13,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(13,3)=Trim(RKS(13,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(13,4)=Trim(RKS(13,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(13,5)=Trim(RKS(13,1))+Trim(A(j))
           Case Else
              RKS(13,6)=Trim(RKS(13,1))+Trim(A(j))
         End  Select      
             
     End  Select      
  Next      

 %>
 
 <!-- </table>-->
 
 </CENTER>
  
 <CENTER>
 <table border-collapse=collapse Border=1 width=80% bgcolor=#FFFF00>
 <TR><TD  width=150></TD><TD width=150></TD><TD width =150></TD><TD width =150></TD></TR>
 <% 
 Dim Row1,Row2,Row3,Row4,Row5,Row6,RowT 
 Row1 ="<TR>" & "<TD width=150>" & GHMCOD(R(6)) & "</TD>"& "<TD width=150>" & GHMCOD(R(7)) & "</TD>"& "<TD width =150>" & GHMCOD(R(8)) & "</TD>"& "<TD width =150>" & GHMCOD(R(9)) & "</TD>"& "</TR>"
 Row2 = "<TR >" & "<TD width=150>" & GHMCOD(R(5)) & "</TD>"& "<TD width=150  RowSpan=2 ColSpan=2>"& GHMCOD(R(13))&Trim(YYNUM)&Trim(MMNUM)&Trim(DDNUM)&Trim(HHNUM)&"<BR>"&"�R�D:"&Trim(MYL)&"<BR>"&"���D:"&Trim(MYB)& "</TD>"& "<TD width =150>" & GHMCOD(R(10)) & "</TD>"& "</TR>"
Row3 = "<TR>" & "<TD width=150>" & GHMCOD(R(4)) & "</TD>"&  "<TD width=150>" & GHMCOD(R(11)) & "</TD>"& "</TR>"
Row4 = "<TR>" & "<TD width=150>" & GHMCOD(R(3)) & "</TD>"& "<TD width=150>" & GHMCOD(R(2)) & "</TD>"& "<TD width =150>" & GHMCOD(R(1)) & "</TD>"& "<TD width =150>" & GHMCOD(R(12)) & "</TD>"& "</TR>"
 Row5 = "<TR>" & "<TD width=150>"  & "</TD>"& "<TD width=150>" & "</TD>"& "<TD width =150>" & "</TD>"& "<TD width =150>" &"</TD>"& "</TR>"
RowT=Row1+Row2+Row3+Row4+Row5
 Response.Write (RowT)
 Row6 = "<TR style='writing-mode:bt-rl' width =150>" & "�K�O����<u>�{���@��:�}���</u>"& "</TR>"

 Response.Write (Row6)

 %>
   <!--&"<tr style='writing-mode:bt-rl'>"&"�{���ۧ@��:<u>�}���</u>,  ����"&"</tr>"
     ''<span style="writing-mode:tb-rl""<table border-collapse=collapse >"&&"</table >"><%=RowT%>�峹���e RowT </span>
 <div   class="titlered"><%=RowT%> </div><span style="writing-mode:tb-rl">�峹���e RowT</span>-->
<% 
 ' '''Response.Write "<TR >"&"<TD>" & GNam2 & "</TD>"
 '   Response.Write   "<TD>" & GNum2 & "</TD>"
 '   Response.Write   "<TD>" & GNum2s(0)&GNum2s(1)&GNum2s(2)&No12 & "</TD>"
 '   Response.Write   "<TD>" & 4*FGAN(3)+2*FGAN(4)+1*FGAN(5) & "</TD>"
 '   Response.Write   "<TD>" & FGAH(5)&FGAH(4)&FGAH(3) & "</TD>" &"</TR>"
  	   
 '  Response.Write "<TR >"&"<TD>" & GNam1 & "</TD>"
 '   Response.Write   "<TD>" & GNum1 & "</TD>"
 '   Response.Write   "<TD>" & GNum1s(0)&GNum1s(1)&GNum1s(2) & "</TD>"
 '   Response.Write   "<TD>" & 4*FGAN(0)+2*FGAN(1)+1*FGAN(2) & "</TD>"
 '   Response.Write  "<TD>" & FGAH(2)&FGAH(1)&FGAH(0) & "</TD>"&"</TR>"
 '   Response.Write "<TR >"&"<TD>" & FGACK & "</TD>"&"</TR>"

  '' Response.Write (div.style="writing-mode=tb-rl")
   
    %>
  
 <!--</TABLE>-->
  
 </CENTER>
 </BODY> 
 
  <%
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '''SQL = "Select * From [A2:B100]"
 ' SQL = "Select * From ���ת�2"
   '''SQL = "Select * From K0_1"
 'Set rls = GetExcelRecordset( "Excel12.xls",  SQL)
 %><!-- <CENTER>
<TABLE Border=1 BGCOLOR=#FFFF00>    
 <TR ><TD width=720></TD></TR>--><!--</TABLE></CENTER>--><%
   
   ''   Row = "<TR>"
 ' Row = "<TR>"
    ''   For i=1 to 75
     '''   Row = Row & "<TD>" & AW(i)& B(i) & "</TD>"
 '   For i=0 to 13
 '       Row = Row & "<TD>" & R(i) & "</TD>"
     ''' Row = Row & "<TD>" & KFE(i) & "</TD>"
          
 '  Next
 '  Response.Write Row & "</TR>"

 
 %><!-- </TABLE></CENTER>--><!-- <TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>--><%
   ''' Row3 = "<TR>"
 '   For i=1 to 13
 '    For J=1 to  6
      '''  For J=1 to  5
   
 '   Response.Write "<TD>" & RKS(i,j) & "</TD>"

 '     Next  
 '  Next
 
 %> 
    <CENTER>
 <TABLE border-collapse=collapse Border=1 width=70% bgcolor=#FFFF00>
 
 <H3><FONT color=#FF0000> �Y�ݭn�R�L���ת�,��<u>Email:tech.t1206@gmail.com</u>, �p���A��,����!</FONT><HR></H3>
  <!--<H2><a href="http://61.222.248.199/HSU-fundb/Login-2r.asp" ><u>�[�J�|��</u></a></h2>-->
 </TABLE> </CENTER>
 <!--<DIVTABLE> <h2><a href="http://class.ruten.com.tw/user/index00.php?s=tang1206">�^����</a></h2>
  <span style="writing-mode:tb-rl">�峹���e <%=RowT%></span>
  </DIVTABLE>  <div style="writing-mode:tb-lr"> �ó]�w�L���峹���e </div><CENTER></CENTER> -->
 
 
<script Language="VB" runat="server">
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
   ''' ���հƵ{��1
   ''''����WORD�إ�
   FUNCTION GHMCOD(Rnm) 
    
   Dim crsTM,crsTMS
   Dim ghern, RowTS, RowTL, RowTT, RowTB
   Dim  RowTN1, RowTN2, RowTN3, RowTN4   
   Dim k
   Dim ghernm
   crsTM =Trim(Rnm)
       ''crsTM =Trim(R(4))
      ' crsTM =Trim(CStr(rs("�D�ؽX")))&".html"
    'crsTM ="math/"&Trim(CStr(rs("�D�ؽX")))&".html"
     ''Response.Write crsTM
      
      crsTMS=SPLIT(crsTM, ",")
 
   For  k=0 to Ubound(crsTMS)
    ' RowTT=RowTT+Trim(crsTMS(k))+"<br>"
    '' RNDSN crsTMs(k),MMNUMS(k)
   ghern = Mid(Trim(crsTMS(k)),1,1)
   ''' If Cint(ghern) <10 Then  
           ghernm=Cstr(Trim(ghern))
   Select Case ghernm
      Case "*","0","1","2","3","4","5","6","7","8","9"
         'RowTN=Trim(crsTMS(k))+","+Trim(RowTN)
       If Mid(Trim(crsTMS(k)),3,1)="(" Then 
        RowTS=Trim(RowTS)+Trim(crsTMS(k))+"<br>"
      else
        RowTL=Trim(RowTL)+Trim(crsTMS(k))+"<br>"
      end if
 
     Case Else
      If k < 7 Then 
        RowTT=Trim(RowTT)+Trim(crsTMS(k))+"<br>"
      else
        RowTB=Trim(RowTB)+Trim(crsTMS(k))+"<br>"
      end if
    End Select   
       
   NEXT
  RowTN1="<table >"& "<tr>"&"<td Border: solid 1px red>"& RowTT &"</td>"& "</tr>"
  RowTN2="<tr>"&"<td >"& RowTB &"</td>"& "</tr>"
  RowTN3="<tr>"& RowTS &"</tr>"
  RowTN4="<tr>"& RowTL &"</tr>"&"</table>" 
  
  'GHMCOD=RowTN1+RowTN2+RowTN3+RowTN4
  'GHMCOD= RowTT + RowTN
  Return (RowTN1+RowTN2+RowTN3+RowTN4)
 
  END FUNCTION
 
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''�p��y�~��g
 
 FUNCTION LGRNDSN(MONG, HHNUMR)
 
 Dim ZGG() = {"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
 Dim ZGN() = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}
 Dim zgtk1 ,zgtk2 ,zgtk3 ,zgtk4 
 Dim n ,m
    For n=0 to 11
        if MONG=ZGG(n) then
           'zgtk1=ZGN(n)
           zgtk1=n
          end if 
      Next   
     For m=0 to 11
        if HHNUMR=ZGG(m) then
           'zgtk2=ZGN(m)
           zgtk2=m
          end if 
     Next   
      zgtk3=zgtk1+zgtk2
      zgtk4=zgtk3 MOD 12
      ' zgtk=ZGG(zgtk4)
        
     '' LGRNDSN=ZGG(zgtk4)
     Return ZGG(zgtk4)
  END FUNCTION 



 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FUNCTION WRNDSN(zzo, zzk)
  Dim Conn As OleDbConnection, Cmd As OleDbCommand
      ''Dim rsdt As OleDbDataReader
      Dim SQL As String, Body As String
      Dim i , n, j As integer
       ''Dim  k=0
     
      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
        '''Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "/HSU-WN/ZUWE1.mdb" )
      Dim Dbpass = "Jet OLEDB:Database Password=tang1206"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()
    Dim rsdt As OleDbDataReader
      ''SQL = "Select * From ���L�� Where ����='" & DDNUM & "'"
      SQL = "Select * From �ҬP����"
      Cmd = New OleDbCommand( SQL, Conn )
      rsdt = Cmd.ExecuteReader()
   'If rsdt Is Nothing Then
       '   Response.Write ("GetExcelRecordset �I�s����!")
       '  Response.End
    'End If 
 
  Dim w(16)
  w(0)="����"  
  Dim wk
  
  ' Part II�G��X��ƪ��u���e�v
  While rsdt.Read()
  
    IF rsdt(0)= zzo  Then
    
      For i=1 to rsdt.FieldCount-1
        if rsdt.GetName(i)=  zzk then
          'if rsdt(i).Name= "�G" then
          ' w(0)="����"
           w(i)=rsdt(i)
           wk =rsdt(i)
         End If 
       Next
     
    End If  
  End While
   rsdt.close()
     Conn.Close()
  
 ''' WRNDSN=wk  
    Return wk
 '' Response.Write (wk)

END FUNCTION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ''�p��Ѥ~�ئ~��
 FUNCTION GGRNDSN(BN3, YKGS)
  Dim  ZGG() = {"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
 Dim  ZGN() = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}
 Dim zgtk1,zgtk2,zgtk3,zgtk4 
 Dim n ,m
     For n=0 to 11
        if BN3=ZGG(n) then
           'zgtk1=ZGN(n)
           zgtk1=n
          end if 
      Next   
     For m=0 to 11
        if YKGS=ZGG(m) then
           'zgtk2=ZGN(m)
           zgtk2=m
          end if 
     Next   
      zgtk3=zgtk1+zgtk2
      zgtk4=zgtk3 MOD 12
      ' zgtk=ZGG(zgtk4)
        ' 'GGRNDSN=ZG(YG)
       'zgtkn=ZGG(zgtk4)
     
     '' GGRNDSN=ZGG(zgtk4)
      Return ZGG(zgtk4)
  'A(0+k)=zgt(0)
  ' B(0+k)=zgtk
 ' AW(0+k)=zgt(0)& wk
  ''AW(0+k)=zgt(0)&wk
  ' A(105)="�Ѥ~��"
  ' B(105)=zgtk
 ' AW(105)=zgt(0)& wk
      
 END FUNCTION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
  ''�p��T�x�K�y�����ѶQ�c��
    
   FUNCTION TGRNDSN(TBNM, TDDNUM)
 
  Dim ZGG() = {"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
  Dim ZGN() = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}
  Dim zmtk1,zmtk2,zmtk3,zmtk4  
  Dim n
    For n=0 to 11
        if TBNM=ZGG(n) then
           'zmtk1=ZGN(n)
           zmtk1=n
           'zmtk2=TDDNUM
         end if 
      Next   
    
       zmtk2=TDDNUM
      zmtk3=zmtk1+zmtk2+36
      zmtk4=zmtk3 MOD 12
        'zmtk=ZGG(zmtk4)
        ''GGRNDSN=ZG(YG)
        TGRNDSN=ZGG(zmtk4)
 
 'A(0+k)=zmt(0)
  ' A(0+k)=�T�x
  ' B(0+k)=zmtk
 ' AW(0+k)=zmt(0)& wk
   
 END FUNCTION 



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
 Dim ZK()={"��","�A","��","�B","��","�v","��","��","��","��"}
 Dim ZG()={"�l","��","�G","�f","��","�w","��","��","��","��","��","��"}
 
 FUNCTION GRNDSN(YYNUM, MMNUM, DDNUM, HHNUM)
  Dim YK,YK8,YG,MG,MG1,DK,DG 
  Dim D1=DateSerial(1912,2,18)
  Dim D2=DateSerial(YYNUM,MMNUM,DDNUM)
  Dim DY=DateDiff("yyyy", D1, D2)
  Dim DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= MMNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
 '' Response.Write ("<TR><TD>���~:"&  D1 & "</TD></TR>")
 '''Response.Write ("<TR><TD>�ͦ~:"&  D2 & "</TD></TR>")
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 '''Response.Write ("<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>")
'''Response.Write( "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"&"<BR>")
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"

  GRNDSN=ZG(YG)
 END FUNCTION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 

 FUNCTION KRNDSN(YYNUM, MMNUM, DDNUM, HHNUM)
  Dim YK,YK8,YG,MG,MG1,DK,DG 
  Dim D1=DateSerial(1912,2,18)
  Dim D2=DateSerial(YYNUM,MMNUM,DDNUM)
  Dim DY=DateDiff("yyyy", D1, D2)
  Dim DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= MMNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
 '' Response.Write ("<TR><TD>���~:"&  D1 & "</TD></TR>")
 '''Response.Write ("<TR><TD>�ͦ~:"&  D2 & "</TD></TR>")
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 '''Response.Write ("<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>")
'''Response.Write( "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"&"<BR>")
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"

  KRNDSN=ZK(YK)
 END FUNCTION 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 </script> 
</HTML>