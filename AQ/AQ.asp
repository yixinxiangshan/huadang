<!--#include file= database.asp-->
<%
if (session("M_StarTime")=0) or (InStr(request.servervariables("http_referer"),"index.asp")>0) then        '��ʼ����
   session("M_DurationTime") = 0       '����ʱ��
   session("M_RightNum") = 0           '��Ե���Ŀ��
   session("M_ErrorNum") = 0           '������Ŀ��
   session("M_StarTime") = timer()     '��ʼʱ��
   session("M_EndTime") = 0            '����ʱ��
   session("lastRequestTimer") = 0     '����ȡ��ʱ��
   session("ydQuestions") = ""         '�Ѿ��������Ŀ
else
   if cdbl(request.form("requestTimer"))<=cdbl(session("lastRequestTimer")) then
      response.Redirect("RefreshOrBackError.asp")
   else
      session("lastRequestTimer") = request.form("requestTimer") 
	  
	  '�Ƿ���ȷ
	  if request.form("aR") = request.form("s90DT") then
	     session("M_RightNum") = session("M_RightNum") + 1
      else
	     session("M_ErrorNum") = session("M_ErrorNum") + 1
	  end if
	  
	  '��������С5����ʧ��
	  if session("M_ErrorNum") > 1 then
	     response.Redirect("fail.asp")
	  end if	
	  
	  if ((session("M_RightNum") + session("M_ErrorNum"))=3) then
	     '����ɹ�
		 session("M_EndTime") = timer()
		 session("M_DurationTime") = request.form("durationTime") 
		 response.Redirect("Success.asp")
	  end if   
	     	 		 	  
   end if  
end if 	
	  
	'ȡ�����
	curSelectInt = RndNumber(498,1)
	do while (InStr(session("ydQuestions"),formatInt3(curSelectInt))>0)
	   curSelectInt = RndNumber(498,1) 
	loop
	  
	session("ydQuestions") = session("ydQuestions") & formatInt3(curSelectInt) & ":"
	
	set rs=server.createobject("adodb.recordset")
	sql = "select * from Questions where ID = " & curSelectInt
    rs.open sql,conn,1,1	
	
	answerList = DanSan()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>2012����ܼ���֪ʶ����</title>
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
body {
	background-color: #387100;
}
</style>
</head>
<script language="Javascript" type="text/javascript"> 

var time_s_0 = <%if request.form("durationTime")="" then response.write("0") else response.write(request.form("durationTime")) end if%>;
 
function number_length(N, L)
{
	N = N + '';
	M_len = L - N.length; 
	for (i = 0; i < M_len; i++)
	{ N = "0" + N;}
	return N;
}

function display_time()
{
	ns = number_length((time_s_0 % 60), 2);
	nm = number_length(Math.floor(time_s_0/60), 2);
	
	document.all.mytime.innerHTML = nm + ':' +  ns;
	document.page_submit.durationTime.value = time_s_0;
	
	time_s_0++;
	if(time_s_0 > 600) 
	{
   		window.location.href = 'fail.asp';
	}
	else 
	    setTimeout("display_time()", 1000);	
}

function submitFun()
{
   selectNum = 0;	
   for(i=0;i <document.page_submit.s90DT.length;i++)  
       if(document.page_submit.s90DT[i].checked)  
          selectNum = i;    
      
   selectNum = selectNum+1;	
	if (String(selectNum)==document.page_submit.aR.value){
		hideElement('divAnswerError');
		showElement('divAnswerRight');
		}
	else{
		showElement('divAnswerError');
		hideElement('divAnswerRight');
		}	
		
	setTimeout('onSubmitForm()', 500);	
}

function onSubmitForm(){
	page_submit.submit()
	}

function showElement(obj) {
   document.getElementById(obj).style.display = "block"
}
function hideElement(obj) {
   document.getElementById(obj).style.display = "none"
}

</SCRIPT>
<body onload='display_time();'> 
<div align="center">
<form name="page_submit" method="post" action="AQ.asp"> 
<table width="960" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
    <!--#include file= include/top.asp-->
    </td>
  </tr>
  <tr>
    <td height="450" align="center" valign="top" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10%"><img src="images/dt_Title.jpg" width="82" height="26"></td>
            <td width="40%">&nbsp;</td>
            <td width="44%" height="35" align="right" valign="middle"><strong class="font14GreenHeight25">����ʱ�䣺</strong></td>
            <td width="6%" align="left" valign="middle"><strong><span id="mytime" class="font14RedHeight"></span>
            </strong></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td align="center"><table width="95%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="40" class="font14GreenHeight25"><strong>��<%=session("M_RightNum") + session("M_ErrorNum") + 1%>��</strong></td>
          </tr>
          <tr>
            <td height="10"></td>
          </tr>
          <tr>
            <td class="font14BlackHeight25"><%=rs("Q_subject")%></td>
          </tr>
          <tr>
            <td height="10"></td>
          </tr>
          <tr>
            <td height="50">
              <span class="font14BlackHeight25">
            <label>
              <input type="radio" name="s90DT" value="1" onClick="submitFun();" />
            </label>
            <%
			  tempInt = Cint(mid(answerList,1,1))
			  Select Case tempInt
			     Case 1
				     response.write(rs("Q_A"))
					 rightAnswer = 1
			     Case 2
				     response.write(rs("Q_B"))		
			     Case 3
				     response.write(rs("Q_C"))				 
			  End Select 
			%></span></td>
          </tr>
          <tr>
            <td height="50"><span class="font14BlackHeight25">
              <input type="radio" name="s90DT" value="2" onClick="submitFun();" />
            <%
			  tempInt = Cint(mid(answerList,2,1))
			  Select Case tempInt
			     Case 1
				     response.write(rs("Q_A"))
					 rightAnswer = 2
			     Case 2
				     response.write(rs("Q_B"))		
			     Case 3
				     response.write(rs("Q_C"))				 
			  End Select 
			%></span></td>
          </tr>
          <tr>
            <td height="50"><span class="font14BlackHeight25">
              <input type="radio" name="s90DT" value="3" onClick="submitFun();" />
            <%
			  tempInt = Cint(mid(answerList,3,1))
			  Select Case tempInt
			     Case 1
				     response.write(rs("Q_A"))
					 rightAnswer = 3
			     Case 2
				     response.write(rs("Q_B"))		
			     Case 3
				     response.write(rs("Q_C"))				 
			  End Select 
			%>
            </span></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="80" align="center" class="font14RedHeight25">
       
       <div id="divAnswerError" style="display:none;"> 
        <table width="200" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="58" align="center"><img src="images/error.gif" width="60" height="60"></td>
            <td width="142" class="font14RedHeight25"><strong>�Բ��𣬴������</strong></td>
          </tr>
        </table>
        </div>
        
        <div id="divAnswerRight" style="display:none;"> 
       <table width="200" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="58"><img src="images/right.gif" width="73" height="80" /></td>
            <td width="142" class="font14GreenHeight25"><strong>��ϲ�㣬�������</strong></td>
          </tr>
       </table>
       <div>
        </td>
      </tr>
      <tr>
        <td height="30" align="center" class="font14RedHeight25"><span class="font14GreenHeight25">��ԣ�<%=session("M_RightNum")%></span>  &nbsp;&nbsp;<span class="font14GreenHeight25">���</span><%=session("M_ErrorNum")%></td>
      </tr>
      <tr>
        <td height="50" align="center"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="15"></td>
  </tr>
  <tr>
    <td align="center">
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" valign="top" class="fontYellowHeight23">��վ��� | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">��������</a> | <a href="/hhdj/bbs" class="yellowNagative">������̳</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">�������</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">��վ��ͼ</a> |<input name="aR" id="aR" value="<%=rightAnswer%>"><a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">��Ȩ����</a><br>
          2012 �й��Ϻ����磨���ţ����޹�˾ίԱ������ ��Ȩ���� <br>
          ������ 1024x768 �ֱ��� </td>
        <td align="right" valign="top" class="fontYellowHeight23">δ����Ȩ��ֹת�ء�ժ�ࡢ���ƻ�������<br>
          Power by Vasisoft</td>
      </tr>
     </table>     
     <input type="hidden" name="durationTime" id="durationTime" value="0">
     <input type="hidden" name="requestTimer" id="requestTimer" value="<%=timer()%>">
    </td>
  </tr>
</table>
</form>
</div>
</body>
</html>
<%closeconn%>
