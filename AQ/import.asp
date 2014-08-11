<!--#include file= database.asp-->
<!--#include file= config.asp-->
<%
	input_array = split(request.Form("input"), vbnewline)
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from Questions"
	rs.open sql,conn,1,3
	
	i = 0
	max_id = 0
	do while i <= ubound(input_array)
		input_array_array = split(input_array(i), "|")
		rs.AddNew
		max_id = max_id + 1
		rs("ID") = max_id
		rs("Q_type") = input_array_array(0)
		rs("Q_subject") = input_array_array(1)
		if InStr(input_array_array(2), "A") > 0 then rs("A_1") = true end if
		if InStr(input_array_array(2), "B") > 0 then rs("A_2") = true end if
		if InStr(input_array_array(2), "C") > 0 then rs("A_3") = true end if
		if InStr(input_array_array(2), "D") > 0 then rs("A_4") = true end if
		if InStr(input_array_array(2), "E") > 0 then rs("A_5") = true end if
		
		if input_array_array(0) = 2 then
			rs("C_1") = ""
			rs("C_2") = ""
			rs("C_3") = ""
		else
			rs("C_1") = input_array_array(3)
			rs("C_2") = input_array_array(4)
			rs("C_3") = input_array_array(5)		
		end if
		if ubound(input_array_array) > 5 then
			rs("C_4") = input_array_array(6)
		else
			rs("C_4") = ""
		end if
		if ubound(input_array_array) = 7 then
			rs("C_5") = input_array_array(7)
		else
			rs("C_5") = ""
		end if
		rs.Update
		i = i + 1
	loop
	
	rs.close
	set rs=nothing	
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014年社会主义核心价值观网上知识竞赛</title>
		<link href="css.css" rel="stylesheet" type="text/css">
		<script type="text/javascript">
		function analyse(){
			var input = document.getElementById("input").value;
			var input_array = input.split((/\r?\n/));
			
			var data = new Array();
	
			data.push('<table border="1" width="100%" bgcolor="#FFFFFF"><tbody>');
			data.push('<tr>');
			data.push('<th width="5%">题目类型</th>');
			data.push('<th width="40%">题目</th>');
			data.push('<th width="5%">答案</th>');
			data.push('<th width="10%">选项A</th>');
			data.push('<th width="10%">选项B</th>');
			data.push('<th width="10%">选项C</th>');
			data.push('<th width="10%">选项D</th>');
			data.push('<th width="10%">选项E</th>');
			data.push('</tr>');
			for (i = 0; i < input_array.length; i++){
				data.push('<tr>');
				var input_array_array = input_array[i].split("|");
				switch (input_array_array[0]) {
					case "1": data.push('<td>普通单选</td>'); break;
					case "2": data.push('<td>普通判断</td>'); break;
					case "3": data.push('<td>普通复选</td>'); break;
					case "4": data.push('<td>补充单选</td>'); break;
					case "5": data.push('<td>补充复选</td>'); break;
					default: break;
				}
									
				data.push('<td>' + input_array_array[1] + '</td>');
				data.push('<td>' + input_array_array[2] + '</td>');
				
				if (input_array_array[0] == "2") {
					data.push('<td></td>');
					data.push('<td></td>');
					data.push('<td></td>');
				} else{				
					data.push('<td>' + input_array_array[3] + '</td>');
					data.push('<td>' + input_array_array[4] + '</td>');
					data.push('<td>' + input_array_array[5] + '</td>');
				}
				
				if (input_array_array.length > 6){
					data.push('<td>' + input_array_array[6] + '</td>');
				}
				else{
					data.push('<td></td>');
				}
				
				if (input_array_array.length > 7){
					data.push('<td>' + input_array_array[7] + '</td>');
				}
				else{
					data.push('<td></td>');
				}
				data.push('</tr>');
			}
			data.push('</tbody><table>');
			document.getElementById('table').innerHTML = data.join('');
			document.getElementById('submit').disabled = false;
		}
		</script>
	</head>
	<body>
		<form name="page_submit" method="post" action="import.asp" ID="page_submit">
		<div align="center">
			<table width="960" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>
						<!--#include file= include/top.asp-->
					</td>
				</tr>
				<tr>
					<td height="400" align="center" margin="15" valign="top" bgcolor="#FFFFFF">
						<textarea name="input" id="input" rows="30" cols="120">1|问题1|A|选项A|选项B|选项C|选项D
2|问题2|B
3|问题3|ABCD|选项A|选项B|选项C|选项D
4|问题4|C|选项A2|选项B2|选项C2|选项D2|选项E2
5|问题5|ABC|选项A|选项B|选项C|选项D
4|问题6|D|选项A2|选项B2|选项C2|选项D2|选项E2
1|问题7|E|选项A|选项B|选项C|选项D|选项E2
5|问题8|ABCDE|选项A2|选项B2|选项C2|选项D2|选项E2</textarea>						
					</td>
				</tr>
				<tr>
					<td height="15" align="center" bgcolor="#ffffff"><input type="button" onclick="analyse();" value="分析" ID="Button1" NAME="Button1"/></td>
				</tr>
				<tr>
					<td><div id="table"></div>						
					</td>
				</tr>				
				<tr>
					<td height="15" align="center" bgcolor="#ffffff">
						<input type="submit" onclick="analyse();" value="确认" ID="submit" NAME="submit" disabled/>
					</td>
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
		</div>
		<map name="Map1">
			<area shape="rect" coords="270,124,336,139" href="index.asp" target="_self">
		</map>
		</form>
	</body>
</html>
