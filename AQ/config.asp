<%

'���ⳬʱ
Function my_timeout()
	my_timeout = 10
End Function

'����ܴ������Ŀ��
Function my_failed_number()
	my_failed_number = 1
End Function

'������Ŀ��
Function my_all_number()
	my_all_number = 3
End Function

'��ҳ��������
Function my_ranking_number()
	my_ranking_number = 20
End Function

Function get_company(company_code)	
	select case company_code
        case "1" 
            get_company = "����NEC"
        case "2" 
            get_company = "����΢����"
        case "3" 
            get_company = "�з�����"
        case "4" 
            get_company = "�����ͨ"							
        case "5" 
            get_company = "���ա�������"
        case "6" 
            get_company = "�ܲ�������Ƽ�"
    end select	
End Function
%>