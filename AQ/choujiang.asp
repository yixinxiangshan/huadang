<!--#include file= database.asp-->
<!--#include file= config.asp-->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<link href="css.css" rel="stylesheet" type="text/css">
		<link href="choujiang.css" rel="stylesheet" type="text/css">
			<script type="text/javascript" src="choujiang.js"></script>
			<script type="text/javascript">
	var m_name = [
				<%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select * from members where M_RightNum = " & my_all_number()
                 rs.open sql,conn,1,1
				 i = 0
				 do while not rs.eof
					name = "'" & Deal(rs("M_Name")) & "(" & Deal(rs("M_MP")) & ")',"
					response.write(name)
					rs.movenext
					i = i +1
                 loop
				%>
	];

	window.ini =[
		{name:"", num:50, clas:"ipt1"}
	];
	window.rep = 0; //�Ƿ���������ظ�����
	window.obo = 0; //һ�γ�һ����0/��1��
	window.arr = []; //�ų������������� ���� [0,1,2,3,4]
	window.v_s = 20;   //������Ӧʱ�䣬��λ���룬����Խ�����Խ��
	window.h_s = 1000;  //������window.obo=1��Ч�����ƽ����Ƭ��ʾʱ�䣬��ʾ�ڼ���̰�ť������
	
	var finished_time = 1;
	function confirm(){
		if (finished_time == window.ini.length){
			var lucky_names = "";			
			var obj = document.getElementsByTagName("span");//�ȵõ����е�SPAN���
			for(var i=0;i<obj.length;i++)
			{
				if(obj[i].className == 'oMain')//�ҳ�span�����class=a���Ǹ����
				{
					var getObj = obj[i];
					value = getObj.innerHTML;//�������innerHTML
					lucky_names += value + "��";
				}
			}
			var form = document.createElement("form");
			form.setAttribute("method", "post");
			form.setAttribute("action", "lucky.asp");

			var hiddenField = document.createElement("input");
			hiddenField.setAttribute("type", "hidden");
			hiddenField.setAttribute("name", "lucky_names");
			hiddenField.setAttribute("value", lucky_names);
			form.appendChild(hiddenField);

			document.body.appendChild(form);
			form.submit();
			
		}
		else{
			finished_time++;
		}
	};
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
						<div class="Top"><div id="Go"><fieldset class="MainBG"><div id="Main"><h2 class="title"><input class="QD" onclick="group();" type="image" src="images/startDraw.jpg" alt="�����齱ϵͳ" <%if i < 50 then response.Write("disabled")%>/></h2>
										<div class="input"><%if i < 50 then response.Write("�����������50�����޷���ʼ�齱")%></div>
									</div>
									<div><input id="start" src="images/startDisabled.jpg" name="start" type="image" alt="��ʼ(�ո�)" disabled />&nbsp;&nbsp;<input id="end" src="images/pauseDisabled.jpg" name="end" type="image" alt="ֹͣ(�ո�)" disabled />&nbsp;&nbsp;<input id="login" src="images/nextRoundDisabled.jpg" name="login" type="image" alt="��һ��(�س�)" onclick="confirm();" disabled /></div>
								</fieldset></div>
							<div id="out"><fieldset><legend> �н���� </legend>
									<ul id="tableOUT">
										<li style="display:none" />
									</ul>
								</fieldset></div>
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
	<%closeconn%>
	</body>
</html>