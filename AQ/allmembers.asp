<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link href="css.css" rel="stylesheet" type="text/css">
			<%			
				set rs=server.createobject("adodb.recordset")
	            sql = "select * from lucky_names"
                rs.open sql,conn,1,1
				lucky_names = rs("lucky_names")
				                    
				Set rs = Server.CreateObject("adodb.recordset")
				sql="select * from members order by M_Company, M_RightNum desc, M_DurationTime"
				rs.open sql,conn,1,1
			%>
	</head>
	<body>
		<div align="center">
			<form name="page_submit" method="post" action="lucky.asp">
				<table width="960" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<!--#include file= include/top.asp-->
						</td>
					</tr>
					<tr>
						<td height="80px" align="center" style="color:red;font-size:50px" bgcolor="#ffffff" vAlign="top">��������������������</td>
					</tr>
					<tr>
						<td>
						<table width="100%" border="0" margin="15px" bgcolor="#ffffff" cellspacing="0" cellpadding="0" ID="Table1">
							<tr class="font14BlackHeight25">								
								<td align="center"><span class="fontBlackHeight23"><strong>����</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>�ɹ���</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>��ʱ</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>������λ</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>�Ƿ��н�</strong></span></td>
							</tr>
						<%
						do while not rs.EOF
							%>
							<tr>								
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=Deal(rs("M_Name"))%></td>
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_DurationTime"))%>��</td>
								<td width="25%" height="24" align="center" class="fontBlackHeight23"><%=get_company(trim(rs("M_Company")))%></td>
								<td width="25%" height="24" align="center" class="fontBlackHeight23"><%if (InStr(lucky_names, rs("M_MP")) > 0) then response.Write("*") end if%></td>
							</tr>
							<%
							rs.MoveNext
						loop
						%>
						</table>
						</td>
					</tr>	
					<tr>
						<td height="40" align="center" bgcolor="#ffffff" vAlign="bottom"> </td>
					</tr>					
					<tr>
						<td width="960" bgcolor="#ffffff" align="center" valign="bottom" colspan="8"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a></td>
					</tr>					
					<tr>
						<td height="20" align="center" bgcolor="#ffffff" vAlign="bottom"> </td>
					</tr>	
					<tr>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td width="50%" valign="top" class="fontYellowHeight23">��վ��� | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">
											��������</a> | <a href="/hhdj/bbs" class="yellowNagative">������̳</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">
											�������</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">
											��վ��ͼ</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">
											��Ȩ����</a><br>
										2014 �й��Ϻ����磨���ţ����޹�˾ίԱ������ ��Ȩ����
										<br>
										������ 1024x768 �ֱ���
									</td>
									<td align="right" valign="top" class="fontYellowHeight23">δ����Ȩ��ֹת�ء�ժ�ࡢ���ƻ�������<br>
										Power by Vasisoft</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</form>
		</div>
		<%closeconn%>
	</body>
</html>