<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link href="css.css" rel="stylesheet" type="text/css">
			<%
				Set rs = Server.CreateObject("adodb.recordset")
				sql="select * from lucky_names"
				rs.open sql,conn,1,3
				if rs.eof then
					rs.addnew
				end if
				if InStr(request.servervariables("http_referer"), "choujiang.asp") > 0 then
					rs("lucky_names") = Deal(request.form("lucky_names"))
					rs.update
				end if
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
						<td height="80px" align="center" style="color:red;font-size:50px" bgcolor="#ffffff" vAlign="top">������</td>
					</tr>
					<tr>
						<td>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table1">
						<%
						names = split(Deal(rs("lucky_names")), "��")
						i = 0
						do while i <> 10
							%>
							<tr>
							<%
							j = 0
							do while j <> 5
								%>
								<td align="center" bgcolor="#ffffff" style="color:red" height="24px" width="170px" ><p><%=names(i * 5 + j)%></p></td>
								<%
								j = j + 1
							loop
							%>
							</tr>
							<%
							i = i + 1
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
