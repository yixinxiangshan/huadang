<!--#include file= database.asp-->
<!--#include file= config.asp-->

<%
'��ʼ��������Ϣ
session("M_DurationTime") = 0    '����ʱ��
session("M_RightNum") = 0        '��Ե���Ŀ��
session("M_ErrorNum") = 0        '������Ŀ��
session("M_StarTime") = 0        '��ʼʱ��
session("M_EndTime") = 0         '����ʱ��
session("lastRequestTimer") = 0  '����ȡ��ʱ��
session("ydQuestions") = ""      '�Ѿ��������Ŀ

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>2014�����������ļ�ֵ������֪ʶ����</title>
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
    <td height="600" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="25">&nbsp;</td>
        <td width="440" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="6"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;����NEC����ӣ�ǰ25����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                  </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                  </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                  </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 25 * from members where M_Company = '1' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=25)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="23" align="center"><%=i%></td>
                    <td height="23" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                      </tr>
                    </table></td>
                    <td height="23" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="23" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                  </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="10"></td>
                </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;����΢���Ӵ���ӣ�ǰ25����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                    </tr>
                  </table></td>
                </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
                </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                    </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                    </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                    </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 25 * from members where M_Company = '2' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=25)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="23" align="center"><%=i%></td>
                    <td height="23" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                      </tr>
                      </table></td>
                    <td height="23" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="23" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                    </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                  </table></td>
                </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
                </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
        <td width="30" valign="top">&nbsp;</td>
        <td width="440" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="6"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;�з����Ĵ���ӣ�ǰ10����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                  </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                  </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                  </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 10 * from members where M_Company = '3' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=10)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="24" align="center"><%=i%></td>
                    <td height="24" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                      </tr>
                    </table></td>
                    <td height="24" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                  </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="5"></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="6"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;�����ͨ����ӣ�ǰ10����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                  </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                  </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                  </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 10 * from members where M_Company = '4' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=10)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="24" align="center"><%=i%></td>
                    <td height="24" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                      </tr>
                    </table></td>
                    <td height="24" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                  </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="11"></td>
          </tr>
          <tr>
            <td height="10"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;���ա����������ӣ�ǰ10����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                  </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                  </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                  </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 10 * from members where M_Company = '5' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=10)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="24" align="center"><%=i%></td>
                    <td height="24" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                      </tr>
                    </table></td>
                    <td height="24" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                  </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                </table></td>
              </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
              </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="10"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="11"></td>
                </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31" /></td>
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;�����ܲ�������Ƽ���չ���ӣ�ǰ10����</strong></td>
                    <td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31" /></td>
                    </tr>
                  </table></td>
                </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
                </tr>
              <tr>
                <td align="center" background="images/Banner_content_bk.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="fontBlackHeight23">
                    <td width="15%" height="24" align="center"><strong>����</strong></td>
                    <td width="30%" height="24" align="center"><strong>����&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>��ȷ</strong></td>
                    <td width="30%" height="24" align="center"><strong>ʱ��(��)</strong></td>
                    </tr>
                  <tr>
                    <td height="1" colspan="4" background="images/Line_Dot.gif"></td>
                    </tr>
                  <tr class="fontBlackHeight23">
                    <td height="8" colspan="4" align="center"></td>
                    </tr>
                  <%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 10 * from members where M_Company = '6' order by M_RightNum DESC,M_DurationTime"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=10)
				%>
                  <tr class="fontBlackHeight23">
                    <td height="24" align="center"><%=i%></td>
                    <td height="24" align="right"><table width="70%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="fontBlackHeight"><%=Deal(rs("M_Name"))%></td>
                        </tr>
                      </table></td>
                    <td height="24" align="center"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>��</td>
                    </tr>
                  <%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
                  </table></td>
                </tr>
              <tr>
                <td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
                </tr>
              <tr>
                <td><img src="images/Banner_bottom.jpg" width="453" height="9" /></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td height="10"></td>
          </tr>
        </table></td>
        <td width="25">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="3" valign="top">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="3" align="center" valign="top"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="3" valign="top">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
   </table></td>
  </tr>
  <tr>
    <td height="15" align="center"></td>
  </tr>
  <tr>
    <td>
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" valign="top" class="fontYellowHeight23">��վ��� | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">��������</a> | <a href="/hhdj/bbs" class="yellowNagative">������̳</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">�������</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">��վ��ͼ</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">��Ȩ����</a><br>
          2014 �й��Ϻ����磨���ţ����޹�˾ίԱ������ ��Ȩ���� <br>
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
<%closeconn%>
<%
Function Deal(exp1)
     dim exp2
     exp2=Replace(exp1," ","")
	 exp2=Replace(exp1,"&nbsp;","")
     exp2=Replace(exp2,"��","")
     Deal=exp2
End Function 
%>