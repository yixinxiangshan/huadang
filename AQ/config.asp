<%
'答题开始时间
Function my_start_date()
	my_start_date = "Mon Sept 1 00:00:00 +0000 2014"
End Function

'答题结束时间
Function my_stop_date()
	my_stop_date = "Fri Sept 26 00:00:00 +0000 2014"
End Function

'答题超时
Function my_timeout()
	my_timeout = 10
End Function

'最大能错误的题目数
Function my_failed_score()
	my_failed_score = 10
End Function

'所有题目数
Function my_all_number()
	my_all_number = my_normal_single_number() + my_normal_multi_number() + my_single_number() + my_multi_number()
End Function

'所有补充单选题目数
Function my_normal_single_number()
	my_normal_single_number = 2
End Function

'所有补充复选题目数
Function my_normal_multi_number()
	my_normal_multi_number = 1
End Function

'所有补充单选题目数
Function my_single_number()
	my_single_number = 2
End Function

'所有补充复选题目数
Function my_multi_number()
	my_multi_number = 1
End Function

'主页排名个数
Function my_ranking_number()
	my_ranking_number = 20
End Function

Function get_company(company_code)	
	select case company_code
        case "1" 
            get_company = "华虹NEC"
        case "2" 
            get_company = "华力微电子"
        case "3" 
            get_company = "研发中心"
        case "4" 
            get_company = "华虹计通"							
        case "5" 
            get_company = "虹日、进出口"
        case "6" 
            get_company = "总部、华虹科技"
    end select	
End Function

Function get_input_type(q_type, input_number)
	select case q_type
		case 1
			get_input_type = "radio"
		case 2
			if (input_number < 3) then
				get_input_type = "radio"
			else
				get_input_type = "hidden"
			end if
		case 3
			get_input_type = "checkbox"
		case 4
			get_input_type = "radio"
		case 5
			get_input_type = "checkbox"
	end select
End Function

Function Deal(exp1)
     dim exp2
     exp2=Replace(exp1," ","")
	 exp2=Replace(exp1,"&nbsp;","")
     exp2=Replace(exp2,"　","")
     Deal=exp2
End Function 
%>