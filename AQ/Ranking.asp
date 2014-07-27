<!--#include file= database.asp-->
<!--#include file= config.asp-->

<%
'初始化基本信息
session("M_DurationTime") = 0    '持续时间
session("M_RightNum") = 0        '答对的题目数
session("M_ErrorNum") = 0        '答错的题目数
session("M_StarTime") = 0        '开始时间
session("M_EndTime") = 0         '结束时间
session("lastRequestTimer") = 0  '最后读取的时间
session("ydQuestions") = ""      '已经答过的题目

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>2014年社会主义核心价值观网上知识竞赛</title>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;华虹NEC代表队（前25名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="23" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;华力微电子代表队（前25名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="23" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;研发中心代表队（前10名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;华虹计通代表队（前10名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;虹日、进出口联队（前10名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
                    <td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;集团总部、华虹科技发展联队（前10名）</strong></td>
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
                    <td width="15%" height="24" align="center"><strong>排行</strong></td>
                    <td width="30%" height="24" align="center"><strong>姓名&nbsp;&nbsp;</strong></td>
                    <td width="30%" height="24" align="center"><strong>正确</strong></td>
                    <td width="30%" height="24" align="center"><strong>时间(秒)</strong></td>
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
                    <td height="24" align="center"><%=cint(rs("M_DurationTime"))%>秒</td>
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
        <td width="50%" valign="top" class="fontYellowHeight23">网站简介 | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">互动交流</a> | <a href="/hhdj/bbs" class="yellowNagative">在线论坛</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">书记信箱</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">网站地图</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">版权声明</a><br>
          2014 中共上海华虹（集团）有限公司委员会主办 版权所有 <br>
          最佳浏览 1024x768 分辨率 </td>
        <td align="right" valign="top" class="fontYellowHeight23">未经授权禁止转载、摘编、复制或建立镜像<br>
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
     exp2=Replace(exp2,"　","")
     Deal=exp2
End Function 
%>