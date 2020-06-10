select 
	st.NAME as 'Станция (для заголовка)',
	LIST_NO as 'Номер ПЛ (для заголовка)',
	CAR_NO as 'Номер вагона', 
	BUILT_YEAR as 'Год постройки', 
	CAR_TYPE as 'Род вагона', 
	CAR_LOCATION as 'Дислокация', 
	ADM_CODE as 'Код страны-собств.', 
	[OWNER] as 'Собственник', 
	CASE 
		WHEN IS_LOADED = 0 THEN 'негруженный' 
		WHEN IS_LOADED = 1 THEN 'груженный' 
	END as 'Состояние', 
	CASE 
		WHEN IS_WORKING = 0 THEN 'нерабочий' 
		WHEN IS_WORKING = 1 THEN 'рабочий' 
		else 'не определено'
	END as 'Парк', 
	CASE 
		WHEN NON_WORKING_STATE = 1 THEN 'неисправный' 
		WHEN NON_WORKING_STATE = 2 THEN 'резерв' 
		WHEN NON_WORKING_STATE = 3 THEN 'ДЛЗО' 
		WHEN NON_WORKING_STATE = 4 THEN 'СТН' 
		WHEN NON_WORKING_STATE = 5 THEN 'Поврежден по акту ВУ-25' 
		else 'не определено'
	END as 'Категория НРП'

from CAR_CENSUS_LISTS ccl 
INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR 
/*WHERE st.ESR = 480009*/ /*Ясиноватая*/
WHERE st.ESR = 480403 /*Донецк*/