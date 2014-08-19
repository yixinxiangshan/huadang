<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<title>2014年社会主义核心价值观网上知识竞赛</title>
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
						<td height="80px" align="center" style="color:red;font-size:50px" bgcolor="#ffffff" vAlign="top">参赛名单（按工会排序）</td>
					</tr>
					<tr>
						<td>
						<table width="100%" border="0" margin="15px" bgcolor="#ffffff" cellspacing="0" cellpadding="0" ID="Table1">
							<tr class="font14BlackHeight25">								
								<td align="center"><span class="fontBlackHeight23"><strong>姓名</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>成功率</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>用时</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>所属单位</strong></span></td>
								<td align="center"><span class="fontBlackHeight23"><strong>是否中奖</strong></span></td>
							</tr>
						<%
						do while not rs.EOF
							%>
							<tr>								
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=Deal(rs("M_Name"))%></td>
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
								<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_DurationTime"))%>秒</td>
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
