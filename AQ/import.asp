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
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<link href="css.css" rel="stylesheet" type="text/css">
		<script type="text/javascript">
		function analyse(){
			var input = document.getElementById("input").value;
			var input_array = input.split((/\r?\n/));
			
			var data = new Array();
	
			data.push('<table border="1" width="100%" bgcolor="#FFFFFF"><tbody>');
			data.push('<tr>');
			data.push('<th width="5%">��Ŀ����</th>');
			data.push('<th width="40%">��Ŀ</th>');
			data.push('<th width="5%">��</th>');
			data.push('<th width="10%">ѡ��A</th>');
			data.push('<th width="10%">ѡ��B</th>');
			data.push('<th width="10%">ѡ��C</th>');
			data.push('<th width="10%">ѡ��D</th>');
			data.push('<th width="10%">ѡ��E</th>');
			data.push('</tr>');
			for (i = 0; i < input_array.length; i++){
				data.push('<tr>');
				var input_array_array = input_array[i].split("|");
				switch (input_array_array[0]) {
					case "1": data.push('<td>��ͨ��ѡ</td>'); break;
					case "2": data.push('<td>��ͨ�ж�</td>'); break;
					case "3": data.push('<td>��ͨ��ѡ</td>'); break;
					case "4": data.push('<td>���䵥ѡ</td>'); break;
					case "5": data.push('<td>���临ѡ</td>'); break;
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
						<textarea name="input" id="input" rows="30" cols="120">1|����1|A|ѡ��A|ѡ��B|ѡ��C|ѡ��D
2|����2|B
3|����3|ABCD|ѡ��A|ѡ��B|ѡ��C|ѡ��D
4|����4|C|ѡ��A2|ѡ��B2|ѡ��C2|ѡ��D2|ѡ��E2
5|����5|ABC|ѡ��A|ѡ��B|ѡ��C|ѡ��D
4|����6|D|ѡ��A2|ѡ��B2|ѡ��C2|ѡ��D2|ѡ��E2
1|����7|E|ѡ��A|ѡ��B|ѡ��C|ѡ��D|ѡ��E2
5|����8|ABCDE|ѡ��A2|ѡ��B2|ѡ��C2|ѡ��D2|ѡ��E2</textarea>						
					</td>
				</tr>
				<tr>
					<td height="15" align="center" bgcolor="#ffffff"><input type="button" onclick="analyse();" value="����" ID="Button1" NAME="Button1"/></td>
				</tr>
				<tr>
					<td><div id="table"></div>						
					</td>
				</tr>				
				<tr>
					<td height="15" align="center" bgcolor="#ffffff">
						<input type="submit" onclick="analyse();" value="ȷ��" ID="submit" NAME="submit" disabled/>
					</td>
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
		</div>
		<map name="Map1">
			<area shape="rect" coords="270,124,336,139" href="index.asp" target="_self">
		</map>
		</form>
	</body>
</html>
