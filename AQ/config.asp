<%

'���ⳬʱ
Function my_timeout()
	my_timeout = 10
End Function

'����ܴ������Ŀ��
Function my_failed_score()
	my_failed_score = 10
End Function

'������Ŀ��
Function my_all_number()
	my_all_number = 6
End Function

'���в��䵥ѡ��Ŀ��
Function my_single_number()
	my_single_number = 2
End Function

'���в��临ѡ��Ŀ��
Function my_multi_number()
	my_multi_number = 1
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
%>