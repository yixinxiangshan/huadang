<%

'答题超时
Function my_timeout()
	my_timeout = 10
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
%>