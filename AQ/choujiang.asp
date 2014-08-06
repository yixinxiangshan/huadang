<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014年社会主义核心价值观网上知识竞赛</title>
		<link href="css.css" rel="stylesheet" type="text/css">
			<script type="text/javascript" src="images/js.js"></script>
			<script type="text/javascript">
	var m_name = [
				<%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select * from members where M_RightNum = " & my_all_number()
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof
					name = "'" & Deal(rs("M_Name")) & ":(" & Deal(rs("M_MP")) & ")',"
					response.write(name)
					rs.movenext
                 loop
				%>
	];

	window.ini =[
		{name:"幸运奖", num:1, clas:"ipt5"}
	];
	window.rep = 0; //是否允许号码重复出现
	window.obo = 0; //一次抽一（组0/个1）
	window.arr = []; //排除号码或姓名序号 例子 [0,1,2,3,4]
	window.v_s = 50;   //滚动响应时间，单位毫秒，数字越大滚动越慢
	window.h_s = 1000;  //本参数window.obo=1有效，控制结果卡片显示时间，显示期间键盘按钮被锁定
			</script>
	</head>
	<body>
		<div align="center">
			<table width="960" border="0" cellspacing="0" cellpadding="0" ID="Table1">
				<tr>
					<td>
						<!--#include file= include/top.asp-->
					</td>
				</tr>
				<tr>
					<td height="400" align="center" valign="top" bgcolor="#FFFFFF">
						<div id="ctrl"></div>
						<div class="oneOut"><div></div>
						</div>
						<div class="Top"><div id="Go"><fieldset class="MainBG"><div id="Main"><h2 class="title"><input class="QD" onclick="group();" type="button" value="启动抽奖系统"/></h2>
										<div class="input"></div>
									</div>
									<div id="input"><input id="start" name="start" type="button" value="开始(空格)" /><input id="end" name="end" type="button" value="停止(空格)" /><input id="login" name="login" type="button" value="下一组(回车)" /></div>
								</fieldset></div>
							<div id="out"><fieldset><legend> 中奖结果 </legend>
									<ul id="tableOUT">
										<li style="display:none" />
									</ul>
								</fieldset></div>
						</div>
						<div class="End" align="center"><ul>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<td height="15" align="center"></td>
				</tr>
				<tr>
					<td>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table2">
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
		</div>
		<map name="Map1">
			<area shape="rect" coords="270,124,336,139" href="index.asp" target="_self">
		</map>
	</body>
</html>
