<!--#include file= database.asp-->
<!--#include file= config.asp-->
<HTML>
	<HEAD>
		<title>2014�����������ļ�ֵ������֪ʶ����</title>
		<%
'��ʼ��������Ϣ
session("M_DurationTime") = 0    '����ʱ��
session("M_RightNum") = 0        '��Ե���Ŀ��
session("M_ErrorNum") = 0        '������Ŀ��
session("M_StarTime") = 0        '��ʼʱ��
session("M_EndTime") = 0         '����ʱ��
session("lastRequestTimer") = 0  '����ȡ��ʱ��
session("ydQuestions") = ""      '�Ѿ��������Ŀ

%>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link href="css.css" rel="stylesheet" type="text/css">
	</HEAD>	
	<body>
		<div align="center">
			<table width="960" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td> <!--#include file= include/top.asp-->
					</td>
				</tr>
				<tr>
					<td height="600" align="center" valign="top" bgcolor="#ffffff"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="25">&nbsp;</td>
								<td width="440" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td height="50"><img src="images/Title_hdsm.jpg" width="85" height="25"></td>
													</tr>
													<tr>
														<td><p class="fontBlackHeight23">
															<strong><span class="fontRedHeight23">�ʱ��</span></strong>��2014��9��1�ա�25��<br>														
															</p>
														</td>
													</tr>
													<tr>
														<td height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
										</tr>
										<tr>
											<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td height="1" background="images/Line_Dot.gif"></td>
													</tr>
													<tr>
														<td height="31" align="center" background="images/answer_Top_Bk.jpg"><table width="95%" border="0" cellspacing="0" cellpadding="0">
																<tr class="fontRedHeight23">
																	<td><span class="fontRedHeight23">���سɹ��˴�ͳ�ƣ�</span>
																		<%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select count(*) as countNum from members"
                 rs.open sql,conn,1,1
				 response.write(rs("countNum"))	
				%>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td align="center"><table width="95%" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td height="75" align="center"><div id="start_div" name="start_div"></div></td>
																</tr>
																<tr>
																	<td height="30" class="fontWhiteHeight23">���������Ա��¼:</td>
																</tr>
																<tr>
																	<td align="center" bgcolor="#f2f2f2"><table width="95%" border="0" cellspacing="0" cellpadding="0">
																			<%
	             set rs=server.createobject("adodb.recordset")
	             sql = "select top 6 * from members order by M_DataTime DESC"
                 rs.open sql,conn,1,1
				 i = 1
				 do while not rs.eof and (i<=6)
				%>
																			<tr>
																				<td width="18%" height="24" align="left" class="fontBlackHeight23"><%=Deal(rs("M_Name"))%></td>
																				<td width="15%" height="24" align="left" class="fontBlackHeight23"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
																				<td width="15%" height="24" align="left" class="fontBlackHeight23"><%=cint(rs("M_DurationTime"))%>��</td>
																				<td width="25%" height="24" align="center" class="fontBlackHeight23"><%=Format_Time(rs("M_DataTime"),2)%></td>
																				<td width="27%" height="24" align="left"><span class="fontBlackHeight23"><%=get_company(trim(rs("M_Company")))%></span></td>
																			</tr>
																			<%                      
				       rs.movenext 
                       i=i+1 
                       loop 
				 %>
																		</table>
																	</td>
																</tr>
																<tr>
																	<td height="8"></td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
										</tr>
										<tr>
											<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td height="1" background="images/Line_Dot.gif"></td>
													</tr>
													<tr>
														<td height="8" valign="middle"></td>
													</tr>
													<tr>
														<td height="38" valign="middle"><img src="images/Title_jsgz.jpg" width="84" height="26"></td>
													</tr>
													<tr>
														<td align="center" valign="middle" class="fontBlackHeight23"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td class="fontBlackHeight23">
																		1) ���ϴ���ÿ30��Ϊһ�飬�ɵ����Զ���ȡ��ÿ�����ʱ��Ϊ60����,��������<br>
																		&nbsp;&nbsp; ����60�����ϻ�úϸ�֤�飬ȫ������߻������֤�顣<br>
																		2) �����������ʵ���ύ������д�������ᡣΪ����ͬ��ͬ���ߣ�����������<br>
																		&nbsp;&nbsp; ��д��ϵ��ʽ���ֻ��ţ���<br>
																		3) ������� ��������֤�������У������ȡ50�����˲����ߣ��������<br>
																		&nbsp;&nbsp; Ʒ1�ݡ�<a href="lucky.asp">�鿴������</a></td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
								<td width="30" valign="top">&nbsp;</td>
								<td width="440" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td height="6"></td>
													</tr>
													<tr>
														<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31"></td>
																	<td width="200" background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;���а�ǰ<%=my_ranking_number()%>����</strong></td>
																	<td align="right" background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23">&nbsp;</td>
																	<td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31"></td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
													</tr>
													<tr>
														<td align="center" background="images/Banner_content_bk.jpg"><table width="92%" border="0" cellpadding="0" cellspacing="0">
																<tr>
																	<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																			<tr>
																				<td height="5"></td>
																			</tr>
																			<tr>
																				<td class="fontBlackHeight23">
																					<table width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table1">
																						<tr class="font14BlackHeight25">
																							<td align="center"><span class="fontBlackHeight23"><strong>����</strong></span></td>
																							<td align="center"><span class="fontBlackHeight23"><strong>����</strong></span></td>
																							<td align="center"><span class="fontBlackHeight23"><strong>�ɹ���</strong></span></td>
																							<td align="center"><span class="fontBlackHeight23"><strong>��ʱ</strong></span></td>
																							<td align="center"><span class="fontBlackHeight23"><strong>������λ</strong></span></td>
																						</tr>
																						<%
	                         set rs = server.createobject("adodb.recordset")
	                         sql = "select top " & my_ranking_number() & " * from members order by M_RightNum DESC,M_DurationTime"
                             rs.open sql,conn,1,1
				             i = 0
				             do while not rs.eof and (i < my_ranking_number())
							 %>
																						<tr>
																							<td width="15%" height="24" align="center" class="fontBlackHeight23"><%=i + 1%></td>
																							<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=Deal(rs("M_Name"))%></td>
																							<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_RightNum")*100/my_all_number())%>%</td>
																							<td width="20%" height="24" align="center" class="fontBlackHeight23"><%=cint(rs("M_DurationTime"))%>��</td>
																							<td width="25%" height="24" align="center" class="fontBlackHeight23"><%=get_company(trim(rs("M_Company")))%></td>
																						</tr>
																						<%
				             rs.movenext 
                             i=i+1 
                             loop					 
						     %>
																					</table>
																				</td>
																			</tr>
																			<tr>
																				<td height="5"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
													</tr>
													<tr>
														<td><img src="images/Banner_bottom.jpg" width="453" height="9"></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td height="10"></td>
										</tr>
										<tr>
											<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td height="6"></td>
													</tr>
													<tr>
														<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="31"><img src="images/Banner_Title_left.jpg" width="31" height="31"></td>
																	<td background="images/Banner_Title_Bk.jpg" class="fontWhiteHeight23Title"><strong>&nbsp;���Ų��������ʱͳ��</strong>&nbsp;</td>
																	<td width="8"><img src="images/Banner_Title_Right.jpg" width="8" height="31"></td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
													</tr>
													<tr>
														<td align="center" valign="top" background="images/Banner_content_bk1.jpg"><table width="98%" border="0" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="35%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																			<tr>
																				<td height="30"></td>
																			</tr>
																			<tr>
																				<td align="center" valign="top"><table width="135" border="0" cellpadding="0" cellspacing="0">
																						<tr>
																							<td height="7" align="right"></td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(1)%></span></td>
																						</tr>
																						<tr>
																							<td height="22" align="right">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(2)%></span></td>
																						</tr>
																						<tr>
																							<td height="22" align="right">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(3)%></span></td>
																						</tr>
																						<tr>
																							<td height="22" align="right">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(4)%></span></td>
																						</tr>
																						<tr>
																							<td height="22" align="right">&nbsp;</td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(5)%></span></td>
																						</tr>
																						<tr>
																							<td height="15" align="right"></td>
																						</tr>
																						<tr>
																							<td align="right"><span class="fontBlackHeight23"><%=get_company(6)%></span></td>
																						</tr>
																					</table>
																				</td>
																			</tr>
																		</table>
																	</td>
																	<td width="65%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																			<tr>
																				<td height="10"></td>
																			</tr>
																			<tr>
																				<td><img src="images/DB_Content_Bk_Top.gif" width="258" height="7"></td>
																			</tr>
																			<tr>
																				<td height="290" valign="top" background="images/DB_Content_Bk_Middle.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="10">&nbsp;</td>
																							<td height="270" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
																									<tr>
																										<td height="20"></td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='1'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
																												</tr>
																											</table>
																										</td>
																									</tr>
																									<tr>
																										<td height="28">&nbsp;</td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='2'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
																												</tr>
																											</table>
																										</td>
																									</tr>
																									<tr>
																										<td height="28">&nbsp;</td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='3'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
																												</tr>
																											</table>
																										</td>
																									</tr>
																									<tr>
																										<td height="24">&nbsp;</td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='4'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
																												</tr>
																											</table>
																										</td>
																									</tr>
																									<tr>
																										<td height="24">&nbsp;</td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='5'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
																												</tr>
																											</table>
																										</td>
																									</tr>
																									<tr>
																										<td height="24">&nbsp;</td>
																									</tr>
																									<tr>
																										<td><table width="20" border="0" cellspacing="0" cellpadding="0">
																												<tr>
																													<%
	                            set rs=server.createobject("adodb.recordset")
	                                sql = "select count(*) as countNum from members where M_Company ='6'"
                                    rs.open sql,conn,1,1
				                    tempNum = cint(rs("countNum"))
				                 %>
																													<td><img src="images/bar_left.gif" width="5" height="16"></td>
																													<td><img src="images/bar_middle.jpg" width="<%=cint((tempNum+10)*0.048)%>" height="16" ></td>
																													<td><img src="images/bar_Right.gif" width="4" height="16"></td>
																													<td>&nbsp;</td>
																													<td class="fontBlackHeight"><%=tempNum%></td>
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
																			<tr>
																				<td><img src="images/DB_Content_Bk_Bottom.gif" width="258" height="9"></td>
																			</tr>
																			<tr>
																				<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="10">&nbsp;</td>
																							<td><img src="images/DB_Content_Bk_SM.gif" width="221" height="44" id="IMG1"></td>
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
													<tr>
														<td height="5" align="center" background="images/Banner_content_bk.jpg"></td>
													</tr>
													<tr>
														<td><img src="images/Banner_bottom.jpg" width="453" height="9"></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td height="10"></td>
										</tr>
									</table>
								</td>
								<td width="25">&nbsp;</td>
							</tr>
						</table>
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
		<%closeconn%>
		<%
Function Deal(exp1)
     dim exp2
     exp2=Replace(exp1," ","")
	 exp2=Replace(exp1,"&nbsp;","")
     exp2=Replace(exp2,"��","")
     Deal=exp2
End Function 
%>
	</body>
</HTML>
<script language="Javascript" type="text/javascript">
	var startDate = parseDate("<%=my_start_date()%>");
	var stopDate = parseDate("<%=my_stop_date()%>");
	var now = new Date();
	if (now < startDate) {
		document.all.start_div.innerHTML = '<a><img src="images/StartAnswer.png" width="162" height="42" border="0"></a>';
	}
	else if (now > stopDate) {
		document.all.start_div.innerHTML = '<a><img src="images/StartAnswer.png" width="162" height="42" border="0"></a>';
	}
	else{
		document.all.start_div.innerHTML = '<a href="AQ.asp"><img src="images/StartAnswer.jpg" width="162" height="42" border="0"></a>';
	}		
	function parseDate(str) {
		var v=str.split(' ');
		return new Date(Date.parse(v[1]+" "+v[2]+", "+v[5]+" "+v[3]+" UTC"));
	} 
</script>
