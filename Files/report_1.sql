select 
	st.NAME as '������� (��� ���������)',
	LIST_NO as '����� �� (��� ���������)',
	CAR_NO as '����� ������', 
	BUILT_YEAR as '��� ���������', 
	CAR_TYPE as '��� ������', 
	CAR_LOCATION as '����������', 
	ADM_CODE as '��� ������-������.', 
	[OWNER] as '�����������', 
	CASE 
		WHEN IS_LOADED = 0 THEN '�����������' 
		WHEN IS_LOADED = 1 THEN '���������' 
	END as '���������', 
	CASE 
		WHEN IS_WORKING = 0 THEN '���������' 
		WHEN IS_WORKING = 1 THEN '�������' 
		else '�� ����������'
	END as '����', 
	CASE 
		WHEN NON_WORKING_STATE = 1 THEN '�����������' 
		WHEN NON_WORKING_STATE = 2 THEN '������' 
		WHEN NON_WORKING_STATE = 3 THEN '����' 
		WHEN NON_WORKING_STATE = 4 THEN '���' 
		WHEN NON_WORKING_STATE = 5 THEN '��������� �� ���� ��-25' 
		else '�� ����������'
	END as '��������� ���'

from CAR_CENSUS_LISTS ccl 
INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR 
/*WHERE st.ESR = 480009*/ /*����������*/
WHERE st.ESR = 480403 /*������*/