<!--#include file= database.asp-->
<%
if (session("M_RightNum")=0 or session("M_RightNum")="") or (session("M_ErrorNum")="") or (session("M_StarTime")=0 or session("M_StarTime")="") or (session("M_EndTime")=0 or session("M_EndTime")="") or (session("ydQuestions")="") or (InStr(request.servervariables("http_referer"),"AQ.asp")=0) then
    response.Redirect("index.asp")    '��ת��ָ��ҳ��
else
    M_DurationTime = session("M_DurationTime")	
end if
%>
<html>
	<head>
		<meta name="vs_showGrid" content="False">
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<link href="css.css" rel="stylesheet" type="text/css">
	</head>
	<script language="javascript">
	function CheckForm()
	{
		if(document.page_success.T_Name.value=="")
		{
			document.all.T_Name_error.innerHTML = '����Ϊ��';
			document.page_success.T_Name.focus();
			return false;
		}
		
		t = document.page_success.T_Name.value;
		if (t != t.replace(/[^\u4E00-\u9FA5]/g,'')){
			document.all.T_Name_error.innerHTML = '�����а����������ַ�';
			document.page_success.T_Name.focus();
			return false;			
			} 	
		
		var len = document.page_success.T_Name.value.match(/[^ -~]/g) == null ? document.page_success.T_Name.value.length : document.page_success.T_Name.value.length + document.page_success.T_Name.value.match(/[^ -~]/g).length 
		if(len > 8)
		{
			document.all.T_Name_error.innerHTML = '�����е㳤(���Ʋ�����4������)';
			document.page_success.T_Name.focus();
			return false;
		}
		
		if(len < 4)
		{
			document.all.T_Name_error.innerHTML = '����̫��(���Ʋ�����2������)';
			document.page_success.T_Name.focus();
			return false;
		}
		
		if(document.page_success.T_Contact.value=="")
		{
			document.all.T_Contact_error.innerHTML = '����Ϊ��';
			document.page_success.T_Contact.focus();
			return false;
		}		
		
		if(isNaN(document.page_success.T_Contact.value)){
			document.all.T_Contact_error.innerHTML = '��ʽ����,�����������ַ�';
			document.page_success.T_Contact.focus();
			return false;			
			} 

		if ((document.page_success.T_Contact.value.length != 11)) 
		{
			document.all.T_Contact_error.innerHTML = '���Ȳ���,��Ϊ11λ�ֻ�';
			document.page_success.T_Contact.focus();
			return false;
		}
		
		if(document.getElementById("T_Company").value==0)
		{
			document.all.T_Company_error.innerHTML = '��ѡ��������˾!';
			document.page_success.T_Company.focus();
			return false;
		}
	}
	
	function clearT_Name_error(){
		document.all.T_Name_error.innerHTML = '';
		}
		
	function clearT_Contact(){
		document.all.T_Contact_error.innerHTML = '';
		}
	
	function clearT_Company(){
		document.all.T_Company_error.innerHTML = '';
		}

</script>
	<body>
		<div align="center">
			<form name="page_success" method="post" action="submit.asp" onSubmit="return CheckForm();">
				<table width="960" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td> <!--#include file= include/top.asp-->
						</td>
					</tr>
					<tr>
						<td height="480" align="center" valign="top" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td width="10%"><img src="images/dt_Title.jpg" width="82" height="26"></td>
												<td width="40%"><strong class="font14RedHeight25">��ϲ������������</strong></td>
												<td width="44%" height="35" align="right" valign="middle"><strong class="font14RedHeight25">����ʱ�䣺</strong></td>
												<td width="6%" align="left" valign="middle"><strong><span id="mytime" class="font14RedHeight">
															<%= right("0" & (cint(M_DurationTime) \ 60),2) & ":" & (cint(M_DurationTime) mod 60) & ((cdbl(session("M_EndTime"))-cdbl(session("M_StarTime")))*100  mod 100)/100%>
														</span></strong></td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td height="80" align="center" class="font14RedHeight25"><div id="divAnswerError" style="display:none;">
											<table width="200" border="0" cellpadding="0" cellspacing="0">
												<tr>
													<td width="58" align="center"><img src="images/error.gif" width="60" height="60"></td>
													<td width="142" class="font14RedHeight25"><strong>�Բ��𣬴������</strong></td>
												</tr>
											</table>
										</div>
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td align="center"><table width="400" border="0" cellspacing="0" cellpadding="0">
														<tr>
															<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																	<tr>
																		<td><img src="images/success_Submit_Top.gif" width="400" height="9"></td>
																	</tr>
																	<tr>
																		<td height="270" align="center" valign="top" background="images/success_Submit_Content_Bk.gif"><table width="90%" border="0" cellspacing="0" cellpadding="0">
																				<tr>
																					<td height="30" colspan="3" align="left" class="blackNagative">&nbsp;</td>
																				</tr>
																				<tr>
																					<td width="29%" align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��    &nbsp;&nbsp;&nbsp;����</span></td>
																					<td width="47%"><span class="font14BlackHeight25">
																							<input name="T_Name" type="text" class="fontBlackHeight" id="T_Name" size="28" onChange="clearT_Name_error()">
																						</span></td>
																					<td width="24%" class="font14RedHeight"><span class="fontRedHeight23">(*����)</span></td>
																				</tr>
																				<tr>
																					<td align="right" class="font14BlackHeight25">&nbsp;</td>
																					<td><span id="T_Name_error" class="fontRedHeight23"></span></td>
																					<td class="font14RedHeight">&nbsp;</td>
																				</tr>
																				<tr>
																					<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��    &nbsp;&nbsp;&nbsp;����</span></td>
																					<td><span class="font14BlackHeight25">
																							<input name="T_Contact" type="text" class="fontBlackHeight" id="T_Contact" size="28" onChange="clearT_Contact()">
																						</span></td>
																					<td class="font14RedHeight"><span class="fontRedHeight23">(*����)</span></td>
																				</tr>
																				<tr>
																					<td align="right" class="font14BlackHeight25">&nbsp;</td>
																					<td><span id="T_Contact_error" class="fontRedHeight23"></span></td>
																					<td class="font14RedHeight">&nbsp;</td>
																				</tr>
																				<tr>
																					<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">�� �� �ӣ�</span></td>
																					<td><select name="T_Company" class="fontBlackHeight" id="T_Company" onChange="clearT_Company()">
																							<option value="0" selected>==��ѡ�����������==</option>
																							<option value="1"><%=get_company(1)%></option>
																							<option value="2"><%=get_company(2)%></option>
																							<option value="3"><%=get_company(3)%></option>
																							<option value="4"><%=get_company(4)%></option>
																							<option value="5"><%=get_company(5)%></option>
																							<option value="6"><%=get_company(6)%></option>
																						</select></td>
																					<td class="fontRedHeight23">(*����)</td>
																				</tr>
																				<tr>
																					<td>&nbsp;</td>
																					<td><span id="T_Company_error" class="fontRedHeight23"></span></td>
																					<td>&nbsp;</td>
																				</tr>
																				<tr>
																					<td height="25">&nbsp;</td>
																					<td>&nbsp;</td>
																					<td>&nbsp;</td>
																				</tr>
																				<tr>
																					<td colspan="3" align="center"><input name="button" type="submit" class="fontBlackHeight" id="button" value="   �ύ   ">
																						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name="button2" type="reset" class="fontBlackHeight" id="button2" value="   ����   "></td>
																				</tr>
																				<tr>
																					<td>&nbsp;</td>
																					<td><input type="hidden" name="H_M_DurationTime" id="H_M_DurationTime" value='<%= session("M_DurationTime") +((cdbl(session("M_EndTime"))-cdbl(session("M_StarTime")))*100  mod 100)/100%>'></td>
																					<td>&nbsp;</td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																	<tr>
																		<td><img src="images/success_Submit_Bottom.gif" width="400" height="9"></td>
																	</tr>
																</table>
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</table>
										<div id="divAnswerRight" style="display:none;">
											<div>
									</td>
								</tr>
								<tr>
									<td height="15" align="center"></td>
								</tr>
								<tr>
									<td height="30" align="center" class="font14RedHeight25">��ԣ�<%=session("M_RightNum")%>
										&nbsp;&nbsp;���<%=session("M_ErrorNum")%></td>
								</tr>
								<tr>
									<td height="50" align="center"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a></td>
								</tr>
							</table>
							<p>&nbsp;</p>
						</td>
					</tr>
					<tr>
						<td height="15" align="center"></td>
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
	</body>
</html>
<%closeconn%>
