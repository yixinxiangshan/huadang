<!--#include file="database.asp"-->
<%
Function Deal(exp1)
     dim exp2
     exp2=Replace(exp1,"<","&lt;")
     exp2=Replace(exp2,">","&gt;")
     exp2=Replace(exp2,"'","''")
     exp2=Replace(exp2,Chr(13),"<br>")
     exp2=Replace(exp2," ","&nbsp;")
     Deal=exp2
End Function 

if (session("M_DurationTime")=0) or (session("M_RightNum")=0 or session("M_RightNum")="") or (session("M_ErrorNum")="") or (session("M_StarTime")=0 or session("M_StarTime")="") or (session("M_EndTime")=0 or session("M_EndTime")="") or (session("ydQuestions")="") or (InStr(request.servervariables("http_referer"),"Success.asp")=0)  then
    response.Redirect("index.asp")    '��ת��ָ��ҳ��
else
    '�Ƿ�����
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql="select * from members where M_Name ='"+ trim(Deal(request.form("T_Name"))) +"'"
	rs.open sql,conn,1,3
	if not rs.eof then 
	   rs.close
	   set rs=nothing
       Response.Write "<script>alert('�Բ���������������Ѵ���!');history.back();</script>"
       response.end
    end if

    '�����¼
	Set rs = Server.CreateObject("ADODB.Recordset")

    '��ȡ���е�����
	M_Name         =       trim(Deal(request.form("T_Name")))
    M_Company      =       request.form("T_Company")
	M_DurationTime =       request.form("H_M_DurationTime")
	M_RightNum     =       session("M_RightNum")
	M_DataTime     =       now()
	M_Num          =       request.form("M_Num")
	M_MP           =       trim(Deal(request.form("T_Contact")))
	M_StarTime     =       session("M_StarTime")
	M_EndTime      =       session("M_EndTime")
	
	sql="select top 1 * from members" 
	rs.open sql,conn,1,3
	rs.addnew	
	   rs("M_Name")  =  M_Name
	   rs("M_Company")  =  M_Company
	   rs("M_DurationTime")  =  M_DurationTime 
	   rs("M_RightNum")  =  M_RightNum
	   rs("M_DataTime")  =  M_DataTime
	   rs("M_Num") =  M_Num
	   rs("M_MP") =  M_MP
	   rs("M_StarTime") =  M_StarTime
	   rs("M_EndTime") =  M_EndTime
	rs.update		
	rs.close
	set rs=nothing
	session("ydQuestions") = ""
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>¡�ؼ����й�����������90����--��ʷ֪ʶ����</title>
<link href="css.css" rel="stylesheet" type="text/css">
</head>
<body>
<div align="center">
<table width="960" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
    <!--#include file= include/top.asp-->
    </td>
  </tr>
  <tr>
    <td height="480" align="center" valign="top" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10%"><img src="images/dt_Title.jpg" width="82" height="26"></td>
            <td width="40%"><strong class="font14RedHeight25">��Ϣ����ȷ�ύ</strong></td>
            <td width="44%" height="35" align="right" valign="middle"><strong class="font14RedHeight">����ʱ�䣺</strong></td>
            <td width="6%" align="left" valign="middle"><strong><span id="mytime" class="font14RedHeight">
             <%= right("0" & (cint(M_DurationTime) \ 60),2) & ":" & (cint(M_DurationTime) mod 60) & ((cdbl(session("M_EndTime"))-cdbl(session("M_StarTime")))*100  mod 100)/100%>
             </span></strong></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="80" align="center" class="font14RedHeight25"><div id="divAnswerError" style="display:none;">
          <table width="200" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="58" align="center"><img src="images/error.gif" width="60" height="60"></td>
              <td width="142" class="font14RedHeight25"><strong>�Բ��𣬴������</strong></td>
              </tr>
            </table>
          </div>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="445"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/success_Submit_Top.gif" width="445" height="9"></td>
                </tr>
                <tr>
                  <td height="270" align="center" valign="top" background="images/success_Submit_Content_Bk.gif"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="40" colspan="3" align="center">&nbsp;</td>
                      </tr>
                    <tr>
                      <td width="19%" align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��    &nbsp;&nbsp;&nbsp;����</span></td>
                      <td width="67%" class="fontRedHeight23"><%=M_Name%></td>
                      <td width="14%" class="font14RedHeight">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="right" class="font14BlackHeight25">&nbsp;</td>
                      <td class="fontRedHeight23">
                      </td>
                      <td class="font14RedHeight">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��ϵ��ʽ��</span></td>
                      <td class="fontRedHeight23"><%=M_MP%></td>
                      <td class="font14RedHeight">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="right" class="font14BlackHeight25">&nbsp;</td>
                      <td class="fontRedHeight23"></td>
                      <td class="font14RedHeight">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��������֯��</span></td>
                      <td class="fontRedHeight23"><%
					  select case trim(M_Company)
                          case "1" 
                                response.write("�Ϻ�����NEC�������޹�˾��ί")
                          case "2" 
                                response.write("�Ϻ�����΢�������޹�˾��ί")
                          case "3" 
                                response.write("�Ϻ����ɵ�·�з��������޹�˾����֧��")
                          case "4" 
                                response.write("�Ϻ����չ��ʵ������޹�˾��֧��")								
                          case "5" 
                                response.write("�Ϻ������ͨ����ϵͳ�ɷ����޹�˾��֧��")
                          case "6" 
                                response.write("�Ϻ����磨���ţ����޹�˾�ܲ���֧��")
                          end select					  
					  %></td>
                      <td class="fontRedHeight23">&nbsp;</td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td ></td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="25">&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="3" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                      </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td><img src="images/success_Submit_Bottom.gif" width="445" height="9"></td>
                </tr>
              </table></td>
              <td width="23">&nbsp;</td>
              <td align="center" valign="top"><table width="406" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="images/ZS_Top.gif" width="406" height="130"></td>
                    </tr>
                    <tr>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="images/ZS_Middle_left.gif" width="189" height="38"></td>
                          <td width="68" align="center" valign="top" bgcolor="#fceab8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="18" align="center" valign="bottom" class="fontBlackZS"><%=M_Name%></td>
                            </tr>
                            <tr>
                              <td><img src="images/ZS_Middle_center_line.gif" width="68" height="2"></td>
                            </tr>
                            <tr>
                              <td height="18" align="center" valign="bottom" class="fontBlackZS"><%=M_Num%></td>
                            </tr>
                          </table></td>
                          <td><img src="images/ZS_Middle_right.gif" width="149" height="38"></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td><img src="images/ZS_Bottom.gif" width="406" height="93"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="28" align="center" bgcolor="#E8E8E8"><a href="javascript:void(0);" class="blackNagative"><strong>����֤��</strong></a></td>
                </tr>
              </table></td>
            </tr>
          </table>
          <div id="divAnswerRight" style="display:none;">
          <div></td>
      </tr>
      <tr>
        <td height="15" align="center"></td>
      </tr>
      <tr>
        <td height="30" align="center" class="font14RedHeight25">��ԣ�<%=session("M_RightNum")%> &nbsp;&nbsp;���<%=session("M_ErrorNum")%> ��ȷ�ʣ�<%=session("M_RightNum")*5 & "%"%></td>
      </tr>
      <tr>
        <td height="50" align="center"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a></td>
      </tr>
    </table>      <p>&nbsp;</p></td>
  </tr>
  <tr>
    <td height="15" align="center"></td>
  </tr>
  <tr>
    <td>
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" valign="top" class="fontYellowHeight23">��վ��� | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">��������</a> | <a href="/hhdj/bbs" class="yellowNagative">������̳</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">�������</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">��վ��ͼ</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">��Ȩ����</a><br>
          2011 �й��Ϻ����磨���ţ����޹�˾ίԱ������ ��Ȩ���� <br>
          ������ 1024x768 �ֱ��� </td>
        <td align="right" valign="top" class="fontYellowHeight23">δ����Ȩ��ֹת�ء�ժ�ࡢ���ƻ�������<br>
          Power by Vasisoft</td>
      </tr>
     </table>
    </td>
  </tr>
</table>
</div>
</body>
</html>