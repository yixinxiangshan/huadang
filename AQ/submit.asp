<!--#include file="database.asp"-->
<%
Function Deal(exp1)
     dim exp2
     exp2=Replace(exp1,"<","&lt;")
     exp2=Replace(exp2,">","&gt;")
     exp2=Replace(exp2,"'","''")
     exp2=Replace(exp2,Chr(13),"<br>")
     exp2=Replace(exp2," ","&nbsp;")
     Deal=exp2
End Function 

if (session("M_DurationTime")=0) or (session("M_RightNum")=0 or session("M_RightNum")="") or (session("M_ErrorNum")="") or (session("M_StarTime")=0 or session("M_StarTime")="") or (session("M_EndTime")=0 or session("M_EndTime")="") or (session("ydQuestions")="") or (InStr(request.servervariables("http_referer"),"Success.asp")=0)  then
    session("M_Name")=""
    session("M_ID")=""
    response.Redirect("index.asp")    '��ת��ָ��ҳ��
else
    '�Ƿ�����
	
    '��ȡ���е�����
	M_Name         =       trim(Deal(request.form("T_Name")))
    M_Company      =       request.form("T_Company")
	M_DurationTime =       request.form("H_M_DurationTime")
	M_RightNum     =       session("M_RightNum")
	M_DataTime     =       now()
	M_Num          =       request.form("M_Num")
	M_MP           =       trim(Deal(request.form("T_Contact")))
	M_StarTime     =       session("M_StarTime")
	M_EndTime      =       session("M_EndTime")	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql="select * from members where M_Name ='"& trim(Deal(request.form("T_Name"))) &"' and M_MP = '" &trim(Deal(request.form("T_Contact"))) &"'"
	rs.open sql,conn,1,3
	if not rs.eof then 
	   isFirst = false
	   
	   oldM_RightNum      = rs("M_RightNum")
	   oldM_DurationTime  = rs("M_DurationTime")
	   oldM_DataTime      = rs("M_DataTime")
	   oldM_Company       = rs("M_Company") 
	   session("M_Name")=rs("M_Name")
	   session("M_ID")=rs("ID")
       
	   if (oldM_RightNum < M_RightNum) or ((oldM_RightNum = M_RightNum) and (CDbl(oldM_DurationTime) > CDbl(M_DurationTime))) then
	   
	      TSMessage = "��ϲ�������γɼ��н�����"
	      rs("M_Name")          =  M_Name
	      rs("M_Company")       =  M_Company
	      rs("M_DurationTime")  =  M_DurationTime 
	      rs("M_RightNum")      =  M_RightNum
	      rs("M_DataTime")      =  M_DataTime
	      rs("M_Num")           =  M_Num
	      rs("M_MP")            =  M_MP
	      rs("M_StarTime")      =  M_StarTime
	      rs("M_EndTime")       =  M_EndTime
	      rs.update		
          rs.close
	      set rs=nothing  		  
		  
	   else
	      TSMessage = "�ɼ������˲������ٽ�������"	
		  rs.close
	      set rs=nothing  	  
	   end if	  
	 
	else
	   isFirst   = true	   
	   TSMessage = ""
	   
	   '��һ�δ���
	   rs.close
	   set rs=nothing
	   	   
       '�����¼
	   Set rs = Server.CreateObject("ADODB.Recordset")
   	   sql="select top 1 * from members" 
	   rs.open sql,conn,1,3
	   rs.addnew	
	      rs("M_Name")          =  M_Name
	      rs("M_Company")       =  M_Company
	      rs("M_DurationTime")  =  M_DurationTime 
	      rs("M_RightNum")      =  M_RightNum
	      rs("M_DataTime")      =  M_DataTime
	      rs("M_Num")           =  M_Num
	      rs("M_MP")            =  M_MP
	      rs("M_StarTime")      =  M_StarTime
	      rs("M_EndTime")       =  M_EndTime
	   rs.update		
       rs.close
	   set rs=nothing  
	   
	   Set rs = Server.CreateObject("ADODB.Recordset")
	   sql2="select top 1 * from members where M_Name ='"& trim(Deal(request.form("T_Name"))) &"' and M_MP = '" &trim(Deal(request.form("T_Contact"))) &"'"
	   rs.open sql2,conn,1,3
	   session("M_ID")=rs("ID")
	   session("M_Name")=rs("M_Name") 
    end if
	 
	session("ydQuestions") = ""
	
    '��ѯһ������
	Set rs = Server.CreateObject("ADODB.Recordset")	
    sql="select count(*) as rankCount from members where ((M_RightNum > " & M_RightNum & ") or ( M_RightNum = " & M_RightNum & " and M_DurationTime <=  "& M_DurationTime  &")) and ( M_Company = '"& M_Company  &"')"
	rs.open sql,conn,1,3
	rankCount = rs("rankCount")
	rs.close
	set rs=nothing	
end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<link href="css.css" rel="stylesheet" type="text/css">
	</head>
	<body>
		<div align="center">
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
											<td width="40%"><strong class="font14RedHeight25">��Ϣ����ȷ�ύ</strong></td>
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
							<%
	  if isFirst then
	  %>
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
																	<td height="40" colspan="3" align="center"></td>
																</tr>
																<tr>
																	<td width="29%" align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��    &nbsp;&nbsp;&nbsp;����</span></td>
																	<td width="47%" class="fontRedHeight23"><%=M_Name%></td>
																	<td width="24%" class="font14RedHeight">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25">&nbsp;</td>
																	<td class="fontRedHeight23"></td>
																	<td class="font14RedHeight">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��ϵ��ʽ��</span></td>
																	<td class="fontRedHeight23"><%=M_MP%></td>
																	<td class="font14RedHeight">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25">&nbsp;</td>
																	<td class="fontRedHeight23"></td>
																	<td class="font14RedHeight">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">������֯��</span></td>
																	<td class="fontRedHeight23"><%=get_company(trim(M_Company))%></td>
																	<td class="fontRedHeight23">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25">&nbsp;</td>
																	<td class="fontRedHeight23"></td>
																	<td class="font14RedHeight">&nbsp;</td>
																</tr>
																<tr>
																	<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">����������</span></td>
																	<td class="fontBlackRankTitle"><%=rankCount%></td>
																	<td>&nbsp;</td>
																</tr>
																<tr>
																	<td colspan="3" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
																</tr>
																<tr>
																	<td>&nbsp;</td>
																	<td>&nbsp;</td>
																	<td>&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td><img src="images/success_Submit_Bottom.gif" width="400" height="8"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<%else%>
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
											<td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr align="center">
														<td height="40" colspan="2" class="font14GreenHeight25"><strong><%=TSMessage%></strong></td>
													</tr>
													<tr align="center">
														<td><table width="400" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																			<tr>
																				<td><img src="images/success_Submit_Top.gif" width="400" height="9"></td>
																			</tr>
																			<tr>
																				<td height="270" align="center" valign="top" background="images/success_Submit_Content_Bk.gif"><table width="90%" border="0" cellspacing="0" cellpadding="0">
																						<tr>
																							<td height="40" colspan="3" align="center"><strong class="fontGreenHeight23">ԭ���ĳɼ�</strong></td>
																						</tr>
																						<tr>
																							<td width="29%" align="right" class="font14BlackHeight25"><span class="fontBlackHeight">���������</span></td>
																							<td width="47%" class="fontRedHeight23"><%=oldM_RightNum %></td>
																							<td width="24%" class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23"></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">����ʱ�䣺</span></td>
																							<td class="fontRedHeight23"><%=cint(oldM_DurationTime)%>��</td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23"></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">������֯��</span></td>
																							<td class="fontRedHeight23"><%=get_company(trim(oldM_Company))%></td>
																							<td class="fontRedHeight23">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23">&nbsp;</td>
																							<td class="fontRedHeight23">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">����ʱ�䣺</span></td>
																							<td class="fontRedHeight23"><%=oldM_DataTime%></td>
																							<td class="fontRedHeight23">&nbsp;</td>
																						</tr>
																						<tr>
																							<td colspan="3" align="center">&nbsp;</td>
																						</tr>
																					</table>
																				</td>
																			</tr>
																			<tr>
																				<td><img src="images/success_Submit_Bottom.gif" width="400" height="8"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
														<td width="445"><table width="400" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																			<tr>
																				<td><img src="images/success_Submit_Top.gif" width="400" height="9"></td>
																			</tr>
																			<tr>
																				<td height="270" align="center" valign="top" background="images/success_Submit_Content_Bk.gif"><table width="90%" border="0" cellspacing="0" cellpadding="0">
																						<tr>
																							<td height="40" colspan="3" align="center" class="fontGreenHeight23"><strong>��ǰ�ĳɼ�</strong></td>
																						</tr>
																						<tr>
																							<td width="29%" align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��    &nbsp;&nbsp;&nbsp;����</span></td>
																							<td width="47%" class="fontRedHeight23"><%=M_Name%></td>
																							<td width="24%" class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23"></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">��ϵ��ʽ��</span></td>
																							<td class="fontRedHeight23"><%=M_MP%></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23"></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">������֯��</span></td>
																							<td class="fontRedHeight23"><%=get_company(trim(M_Company))%></td>
																							<td class="fontRedHeight23">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25">&nbsp;</td>
																							<td class="fontRedHeight23"></td>
																							<td class="font14RedHeight">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right" class="font14BlackHeight25"><span class="fontBlackHeight">����������</span></td>
																							<td class="fontBlackRankTitle"><%=rankCount%></td>
																							<td>&nbsp;</td>
																						</tr>
																						<tr>
																							<td colspan="3" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
																						</tr>
																					</table>
																				</td>
																			</tr>
																			<tr>
																				<td><img src="images/success_Submit_Bottom.gif" width="400" height="8"></td>
																			</tr>
																		</table>
																	</td>
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
							<%end if%>
							<tr>
								<td height="15" align="center"></td>
							</tr>
							<tr>
								<td height="30" align="center" class="font14RedHeight25">��ԣ�<%=session("M_RightNum")%>
									&nbsp;&nbsp;���<%=session("M_ErrorNum")%></td>
							</tr>
							<tr>
								<td height="50" align="center"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a>&nbsp;<a href="cert.asp"><img src="images/printCert.jpg" width="139" height="40" border="0"></a></td>
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
		</div>
	</body>
</html>
