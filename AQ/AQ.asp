<!--#include file= database.asp-->
<!--#include file= config.asp-->
<%
if (session("M_StarTime")=0) and (InStr(request.servervariables("http_referer"),"index.asp")>0) then        '开始答题
   session("M_DurationTime") = 0       '持续时间
   session("M_RightNum") = 0           '答对的题目数
   session("M_ErrorNum") = 0           '答错的题目数
   session("M_ErrorScore") = 0         '答错的题目分数
   session("M_StarTime") = timer()     '开始时间
   session("M_EndTime") = 0            '结束时间
   session("lastRequestTimer") = 0     '最后读取的时间
   session("ydQuestions") = ""         '已经答过的题目
   session("ydNormalSingleNumber") = 0 '已经答过的普通单选题目数
   session("ydNormalMultiNumber") = 0  '已经答过的普通复选题目数
   session("ydSingleNumber") = 0       '已经答过的补充单选题目数
   session("ydMultiNumber") = 0        '已经答过的补充复选题目数
else
   if cdbl(request.form("requestTimer"))<=cdbl(session("lastRequestTimer")) then
      response.Redirect("RefreshOrBackError.asp")
   else
      session("lastRequestTimer") = request.form("requestTimer") 
	  
	  '是否正确
	  if request.form("isCorrect") then
	     session("M_RightNum") = session("M_RightNum") + 1
      else
	     session("M_ErrorNum") = session("M_ErrorNum") + 1
	     session("M_ErrorScore") = session("M_ErrorScore") + request.form("questionScore")
	  end if
	  
	  '如果错误分数大小41，则失败	  
	  if session("M_ErrorScore") > my_failed_score() then
	     response.Redirect("fail.asp")
	  end if
	  
	  if ((session("M_RightNum") + session("M_ErrorNum")) = my_all_number()) then
	     '答题成功
		 session("M_EndTime") = timer()
		 session("M_DurationTime") = request.form("durationTime") 
		 response.Redirect("Success.asp")
	  end if   
	     	 		 	  
   end if  
end if 	
	  
	'取随机数
	set rs=server.createobject("adodb.recordset")
	sql = "select count(*) as countNum from Questions"
    rs.open sql,conn,1,1
	tempNum = cint(rs("countNum"))
	
	curSelectInt = RndNumber(tempNum,1)
	signal = true
	do while (signal)
	   signal = false
	   curSelectInt = RndNumber(tempNum,1) 
	   if InStr(session("ydQuestions"), formatInt3(curSelectInt)) > 0 then
		  signal = true
	   else
	      set rs=server.createobject("adodb.recordset")
		  sql = "select * from Questions where ID = " & curSelectInt
		  rs.open sql,conn,1,1
		  if (rs("Q_type") = 1 Or rs("Q_type") = 2) And session("ydNormalSingleNumber") = my_normal_single_number() then signal = true end if
		  if rs("Q_type") = 3 And session("ydNormalMultiNumber") = my_normal_multi_number() then signal = true end if
		  if rs("Q_type") = 4 And session("ydSingleNumber") = my_single_number() then signal = true end if
		  if rs("Q_type") = 5 And session("ydMultiNumber") = my_multi_number() then signal = true end if
	   end if
	loop
	
	session("ydQuestions") = session("ydQuestions") & formatInt3(curSelectInt) & ":"
	
	set rs=server.createobject("adodb.recordset")
	sql = "select * from Questions where ID = " & curSelectInt
    rs.open sql,conn,1,1	
        
	answerList = DanSan()
	
	rightAnswer = ""
	answer = mid(answerList,1,1)
    if (rs("A_" & answer)) then
		rightAnswer = rightAnswer & ":1"
	end if
	answer = mid(answerList,2,1)
    if (rs("A_" & answer)) then
		rightAnswer = rightAnswer & ":2"
	end if
	answer = mid(answerList,3,1)
    if (rs("A_" & answer)) then
		rightAnswer = rightAnswer & ":3"
	end if
	answer = mid(answerList,4,1)
    if (rs("A_" & answer)) then
		if rs("C_4") <> "" then
			rightAnswer = rightAnswer & ":4"
		else
			rightAnswer = rightAnswer & ":" & InStr(answerList, "4") 
		end if
	end if
	
	'add question score
	if rs("Q_type") = 3 Or rs("Q_type") = 5 then
		questionScore = 4
	else
		questionScore = 3
	end if
	
	'add question number
	if rs("Q_type") = 1 Or rs("Q_type") = 2 then
		session("ydNormalSingleNumber") = session("ydNormalSingleNumber") + 1
	end if
	if rs("Q_type") = 3 then
		session("ydNormalMultiNumber") = session("ydNormalMultiNumber") + 1
	end if
	if rs("Q_type") = 4 then
		session("ydSingleNumber") = session("ydSingleNumber") + 1
	end if
	if rs("Q_type") = 5 then
		session("ydMultiNumber") = session("ydMultiNumber") + 1
	end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
		<title>2014年社会主义核心价值观网上知识竞赛</title>
		<link href="css.css" rel="stylesheet" type="text/css">
	</head>
	<script language="Javascript" type="text/javascript"> 
window.history.forward(1);
var time_s_0 = <%if request.form("durationTime")="" then response.write("0") else response.write(request.form("durationTime")) end if%>;
var this_time = 0;
 
function number_length(N, L)
{
	N = N + '';
	M_len = L - N.length; 
	for (i = 0; i < M_len; i++)
	{ N = "0" + N;}
	return N;
}

function display_time()
{
	ns = number_length((time_s_0 % 60), 2);
	nm = number_length(Math.floor(time_s_0/60), 2);
	
	document.all.mytime.innerHTML = nm + ':' +  ns;
	document.page_submit.durationTime.value = time_s_0;
	
	time_s_0++;
	
	this_ns = number_length((this_time % 60), 2);
	this_nm = number_length(Math.floor(this_time/60), 2);
	
	document.all.thistime.innerHTML = this_nm + ':' +  this_ns;
	
	this_time++;
	if(this_time > <%=my_timeout()%>) 
	{
   		window.location.href = 'timeout.asp';
	}
	else {
	    setTimeout("display_time()", 1000);	
	}
}

function submitFun()
{
   isCorrect = true;
   for(i=0;i <document.page_submit.s90DT.length;i++)  {
      if(document.page_submit.s90DT[i].checked) {          
         if (document.page_submit.aR.value.indexOf(String(i + 1)) == -1) {
	        isCorrect = false;
	        break;
         }
      }
      else {
         if (document.page_submit.aR.value.indexOf(String(i + 1)) != -1) {
	        isCorrect = false;
	        break;
         }
      }
   }
      
	if (isCorrect){
		hideElement('divAnswerError');
		showElement('divAnswerRight');
		}
	else{
		showElement('divAnswerError');
		hideElement('divAnswerRight');
		}	
	document.page_submit.isCorrect.value = isCorrect;
	setTimeout('onSubmitForm()', 500);	
}

function onSubmitForm(){
	page_submit.submit()
	}

function showElement(obj) {
   document.getElementById(obj).style.display = "block"
}
function hideElement(obj) {
   document.getElementById(obj).style.display = "none"
}

	</script>
	<body onload='display_time();'>
		<div align="center">
			<form name="page_submit" method="post" action="AQ.asp">
				<table width="960" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<!--#include file= include/top.asp-->
						</td>
					</tr>
					<tr>
						<td height="450" align="center" valign="top" bgcolor="#FFFFFF"><table width="900" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td width="10%"><img src="images/dt_Title.jpg" width="82" height="26"></td>
												<td width="50%">&nbsp;</td>
												<td width="14%" height="35" align="right" valign="middle"><strong class="font14RedHeight25">总答题时间：</strong></td>
												<td width="6%" align="left" valign="middle"><strong><span id="mytime" class="font14RedHeight"></span></strong>
												<td width="14%" height="35" align="right" valign="middle"><strong class="font14RedHeight25">本题时间：</strong></td>
												<td width="6%" align="left" valign="middle"><strong><span id="thistime" class="font14RedHeight"></span></strong>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td align="center"><table width="95%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td height="40" class="font14RedHeight25"><strong>第<%=session("M_RightNum") + session("M_ErrorNum") + 1%>题</strong></td>
											</tr>
											<tr>
												<td height="10"></td>
											</tr>
											<tr>
												<td class="font14BlackHeight25"><%=rs("Q_subject")%></td>
											</tr>
											<tr>
												<td height="10"></td>
											</tr>
											<tr>
												<td height="30">
													<span class="font14BlackHeight25">
														<label>
															<input type='<%=get_input_type(rs("Q_type"), 1)%>' name="s90DT" value="1" />
														</label>
														<%
              if rs("Q_type") = 2 then
				response.write("对")
				if rs("A_1") then
				  rightAnswer = ":1"
				end if
			  else
				tempInt = Cint(mid(answerList,1,1))
				Select Case tempInt
					Case 1
						response.write(rs("C_1"))
					Case 2
						response.write(rs("C_2"))		
					Case 3
						response.write(rs("C_3"))				
					Case 4
						if rs("C_4") <> "" then							
							response.write(rs("C_4"))
						else
							tempInt = Cint(mid(answerList,4,1))
							Select Case tempInt
								Case 1
									response.write(rs("C_1"))
								Case 2
									response.write(rs("C_2"))		
								Case 3
									response.write(rs("C_3"))
							End Select
						end if
				End Select
			  end if 
			%>
													</span></td>
											</tr>
											<tr>
												<td height="30"><span class="font14BlackHeight25">
														<input type='<%=get_input_type(rs("Q_type"), 2)%>' name="s90DT" value="2" />
														<%
              if rs("Q_type") = 2 then
				response.write("错")
				if rs("A_2") then
				  rightAnswer = ":2"
				end if
			  else
				tempInt = Cint(mid(answerList,2,1))
				Select Case tempInt
					Case 1
						response.write(rs("C_1"))
					Case 2
						response.write(rs("C_2"))		
					Case 3
						response.write(rs("C_3"))				 
					Case 4
						if rs("C_4") <> "" then							
							response.write(rs("C_4"))
						else
							tempInt = Cint(mid(answerList,4,1))
							Select Case tempInt
								Case 1
									response.write(rs("C_1"))
								Case 2
									response.write(rs("C_2"))		
								Case 3
									response.write(rs("C_3"))
							End Select
						end if
				End Select 
			  end if
			%>
													</span></td>
											</tr>
											<tr>
												<td height="30"><span class="font14BlackHeight25">
														<input type='<%=get_input_type(rs("Q_type"), 3)%>' name="s90DT" value="3" />
														<%
            if rs("Q_type") <> 2 then
				tempInt = Cint(mid(answerList,3,1))
				Select Case tempInt
					Case 1
						response.write(rs("C_1"))
					Case 2
						response.write(rs("C_2"))		
					Case 3
						response.write(rs("C_3"))				 
					Case 4
						if rs("C_4") <> "" then							
							response.write(rs("C_4"))
						else
							tempInt = Cint(mid(answerList,4,1))
							Select Case tempInt
								Case 1
									response.write(rs("C_1"))
								Case 2
									response.write(rs("C_2"))		
								Case 3
									response.write(rs("C_3"))
							End Select
						end if
				End Select 
			end if
			%>
													</span></td>
											</tr>
											<tr>
												<td height="30"><span class="font14BlackHeight25">
														<input type='<%if rs("C_4") <> "" then response.Write(get_input_type(rs("Q_type"), 4)) else response.Write("hidden") end if%>' name="s90DT" value="4" />
														<%
			if rs("C_4") <> "" then
				if rs("Q_type") <> 2 then
				tempInt = Cint(mid(answerList,4,1))
				Select Case tempInt
					Case 1
						response.write(rs("C_1"))
					Case 2
						response.write(rs("C_2"))		
					Case 3
						response.write(rs("C_3"))				 
					Case 4
						response.write(rs("C_4"))				 
				End Select 
				end if
			end if
			%>
													</span></td>
											</tr>
											<tr>
												<td height="30"><span class="font14BlackHeight25">
														<input type='<%if rs("C_5") <> "" then response.Write(get_input_type(rs("Q_type"), 5)) else response.Write("hidden") end if%>' name="s90DT" value="5" />
														<%
            if rs("C_5") <> "" then
				     response.write(rs("C_5"))
				     if rs("A_5") then rightAnswer = rightAnswer & ":5"	end if
			end if
			%>
													</span></td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td height="60" align="center" class="font14RedHeight25">
										<div id="divAnswerError" style="display:none;">
											<table width="200" border="0" cellpadding="0" cellspacing="0">
												<tr>
													<td width="58" align="center"><img src="images/error.gif" width="60" height="60"></td>
													<td width="142" class="font14RedHeight25"><strong>对不起，答错啦！</strong></td>
												</tr>
											</table>
										</div>
										<div id="divAnswerRight" style="display:none;">
											<table width="200" border="0" cellpadding="0" cellspacing="0">
												<tr>
													<td width="58"><img src="images/right.gif" width="60" height="60" /></td>
													<td width="142" class="font14RedHeight25"><strong>恭喜你，答对啦！</strong></td>
												</tr>
											</table>
											<div>
									</td>
								</tr>
								<tr>
									<td height="30" align="center" class="font14RedHeight25"><span class="font14RedHeight25">答对：<%=session("M_RightNum")%>题</span>
										&nbsp;&nbsp;<span class="font14RedHeight25">答错：</span><%=session("M_ErrorNum")%>题&nbsp;&nbsp;<span class="font14RedHeight25">已扣分数：</span><%=session("M_ErrorScore")%>分</td>
								</tr>
								<tr>
									<td height="40" align="center"><a href="index.asp"><img src="images/back.jpg" width="139" height="40" border="0"></a>
										<a onClick="submitFun();"><img src="images/next.jpg" width="139" height="40" border="0"></a></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td height="15"></td>
					</tr>
					<tr>
						<td align="center">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td width="50%" valign="top" class="fontYellowHeight23">网站简介 | <a href="/hhdj/webpage/comment.asp" target="_blank" class="yellowNagative">
											互动交流</a> | <a href="/hhdj/bbs" class="yellowNagative">在线论坛</a> | <a href="mailto:shuji@huahong.com.cn" class="yellowNagative">
											书记信箱</a> | <a href="/hhdj/webpage/sitemap.htm" target="_blank" class="yellowNagative">
											网站地图</a> |<input name="aR" type="hidden" id="aR" value="<%=rightAnswer%>"><a href="/hhdj/webpage/copyright.htm" target="_blank" class="yellowNagative">版权声明</a><br>
										2014 中共上海华虹（集团）有限公司委员会主办 版权所有
										<br>
										最佳浏览 1024x768 分辨率
									</td>
									<td align="right" valign="top" class="fontYellowHeight23">未经授权禁止转载、摘编、复制或建立镜像<br>
										Power by Vasisoft</td>
								</tr>
							</table>
							<input type="hidden" name="durationTime" id="durationTime" value="0"> <input type="hidden" name="requestTimer" id="requestTimer" value="<%=timer()%>">
							<input type="hidden" name="isCorrect" id="isCorrect" value=""> <input type="hidden" name="questionScore" id="questionScore" value="<%=questionScore%>">
						</td>
					</tr>
				</table>
			</form>
		</div>
	</body>
</html>
<%closeconn%>
