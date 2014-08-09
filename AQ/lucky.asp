<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<title>2014年社会主义核心价值观网上知识竞赛</title>
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
						<td height="80px" align="center" style="color:red;font-size:50px" bgcolor="#ffffff" vAlign="top">获奖名单</td>
					</tr>
					<tr>
						<td>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table1">
						<%
						names = split(Deal(rs("lucky_names")), "；")
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
									<td width="50%" valign="top" class="fontYellowHeight23">网站简介 | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">
											互动交流</a> | <a href="/hhdj/bbs" class="yellowNagative">在线论坛</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">
											书记信箱</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">
											网站地图</a> | <a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">
											版权声明</a><br>
										2014 中共上海华虹（集团）有限公司委员会主办 版权所有
										<br>
										最佳浏览 1024x768 分辨率
									</td>
									<td align="right" valign="top" class="fontYellowHeight23">未经授权禁止转载、摘编、复制或建立镜像<br>
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
