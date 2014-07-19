<%
dim conn
dim aconnstring
dim DBPath

set conn=server.createobject("adodb.connection")
DBPath = server.mappath("DB/JNJPContest.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

sub closeconn()    
    conn.close
    set conn=nothing
end sub

Function DateFormat8(exp1)
     dim tempContent,returnStr	 
     tempContent = split(exp1,"-")
	 if ubound(tempContent) = 2 then
	    returnStr = tempContent(0) & "-"
		if len(tempContent(1)) = 1 then
		   returnStr = returnStr & "0" & tempContent(1) & "-"
		else
		   returnStr = returnStr & tempContent(1) & "-"
		end if  
		if len(tempContent(2)) = 1 then
		   returnStr = returnStr & "0" & tempContent(2)
		else
		   returnStr = returnStr & tempContent(2)
		end if  
		DateFormat8 = returnStr
	 else
	    DateFormat8 = exp1
	 end if
End Function 

Function RndNumber(MaxNum,MinNum)
    Randomize 
    RndNumber=int((MaxNum-MinNum+1)*rnd+MinNum)
    RndNumber=RndNumber
End Function

Function formatInt3(num)
   tempStr = Cstr(num)
   if (len(tempStr)>3) then
      formatInt3 = right(tempStr,3)
   elseif (len(tempStr)=3) then
      formatInt3 = tempStr
   else
      formatInt3 = right("000"&tempStr,3)	
   end if      	  
End Function

Function DanSan()
   tempStr = "123"
   n1 = RndNumber(3,1) 
      
   returnStr=n1
   tempStr = mid(tempStr,1,n1-1) & mid(tempStr,n1+1,3-n1) 	 
   
   n1 = RndNumber(2,1) 
   returnStr = returnStr& mid(tempStr,n1,1)
   
   tempStr = mid(tempStr,1,n1-1) & mid(tempStr,n1+1,2-n1)
   
   returnStr = returnStr & tempStr
   
   DanSan = returnStr
End Function

' 格式化日期时间(显示) 
' 参数：n_Flag 
' 1:"yyyy-mm-dd hh:mm:ss" 
' 2:"yyyy-mm-dd" 
' 3:"hh:mm:ss" 
' 4:"yyyy年mm月dd日" 
' 5:"yyyymmdd" 
' 6:"yyyymmddhhmmss" 
' 7:"yy-mm-dd"  
' 8:"yy-mm-dd hh:mm:ss" 
' 9:"yyyy年mm月" 
' 10:"mm/dd/yyyy" 
' ============================================ 
Function Format_Time(s_Time, n_Flag) 
   Dim y, m, d, h, mi, s 
       Format_Time = "" 
      If IsDate(s_Time) = False Then Exit Function 
         y = cstr(year(s_Time)) 
         if y = "1900" then Exit Function 
         m = right("0"&month(s_Time),2) 
         d = right("0"&day(s_Time),2) 
         h = right("0"&hour(s_Time),2) 
         mi = right("0"&minute(s_Time),2) 
         s = right("0"&second(s_Time),2) 

         Select Case n_Flag 
         Case 1 
                  Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
         Case 2 
                  Format_Time = y & "-" & m & "-" & d 
         Case 3 
                  Format_Time = h & ":" & mi & ":" & s 
         Case 4 
                  Format_Time = y & "年" & m & "月" & d & "日" 
         Case 5 
                  Format_Time = y & m & d 
         case 6 
                  Format_Time= y & m & d & h & mi & s 
         case 7 
                  Format_Time= right(y,2) & "-" & m & "-" & d 
         case 8 
                  Format_Time= right(y,2) & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
         Case 9 
                  Format_Time = y & "年" & m & "月" 
         Case 10 
                  Format_Time = m & "/" & d & "/" & y & "/" 
         End Select 
End Function 

%>