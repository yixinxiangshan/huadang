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
				rs("lucky_names") = Deal(request.form("lucky_names"))
				rs.update
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
						<td height="400" align="center" bgcolor="#ffffff" vAlign="top"><p><%=Deal(request.form("lucky_names"))%></p>
						</td>
					</tr>
					<tr>
						<td align="center" bgcolor="#ffffff"><input type="hidden" value="清除幸运人员名单" ID="Submit1" NAME="Submit1"></td>
					</tr>
					<tr>
						<td height="15" align="center"></td>
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
