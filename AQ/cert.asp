<%
Dim Is_Excellent 

Is_Excellent=False 

if session("M_RightNum")>0 and session("M_ErrorNum")=0 then
	Is_Excellent=True
end if 
	 
%>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>2014�����������ļ�ֵ������֪ʶ����</title>
	<link href="css.css" rel="stylesheet" type="text/css">
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
				<td align="center" bgcolor="#FFFFFF">
					<!-- Save for Web Slices (cert.jpg) -->
					<table id="Table_01" align="center" width="782" height="502" border="0" cellpadding="0"
						cellspacing="0">
						<tr>
							<td colspan="5">
								<img src="images/cert_01.gif" width="782" height="199" alt=""></td>
						</tr>
						<tr>
							<td>
								<img src="images/cert_02.gif" width="312" height="40" alt=""></td>
							<td colspan="2">
							<% if Is_Excellent then %>
								<img src="images/cert2_03.gif" width="81" height="40" alt=""></td>
							<% else %>
								<img src="images/cert_03.gif" width="81" height="40" alt=""></td>
							<% end if %>
				<td colspan="2">
					<img src="images/cert_04.gif" width="389" height="40" alt=""></td>
			</tr>
			<tr>
				<td colspan="5">
					<img src="images/cert_05.gif" width="782" height="34" alt=""></td>
			</tr>
			<tr>
				<td colspan="2">
					<img src="images/cert_06.gif" width="370" height="40" alt=""></td>
				<td colspan="2" background="images/cert_07.gif" width="125" height="40" alt="" align="center">
					<strong class="font20BlackHeight25">
						<%=session("M_Name")%>
					</strong></span></td>
				<td>
					<img src="images/cert_08.gif" width="287" height="40" alt=""></td>
			</tr>
			<tr>
				<td colspan="2">
					<img src="images/cert_09.gif" width="370" height="37" alt=""></td>
				<td colspan="2" background="images/cert_10.gif" width="125" height="37" alt="" align="center">
					<strong class="font20BlackHeight25">
						<%= Right(String(10, "0") & session("M_ID"), 5)%>
					</strong>
				</td>
				<td>
					<img src="images/cert_11.gif" width="287" height="37" alt=""></td>
			</tr>
			<tr>
				<td colspan="5">
					<img src="images/cert_12.gif" width="782" height="151" alt=""></td>
			</tr>
			<tr>
				<td>
					<img src="images/spacer.gif" width="312" height="1" alt=""></td>
				<td>
					<img src="images/spacer.gif" width="58" height="1" alt=""></td>
				<td>
					<img src="images/spacer.gif" width="23" height="1" alt=""></td>
				<td>
					<img src="images/spacer.gif" width="102" height="1" alt=""></td>
				<td>
					<img src="images/spacer.gif" width="287" height="1" alt=""></td>
			</tr>
			<tr>
				<td align="center" colspan="8"><a href="index.asp"> <img src="images/back.jpg" width="139" height="40" border="0"></a>
				</td>
			</tr>
		</table>
		<!-- End Save for Web Slices -->
		</td> </tr>
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
</body>
</html>
