
#Область СлужебныйПрограммныйИнтерфейс

#Область ВыгрузкаШаблоновПроформВExcel

// Функция конструирует имя листа эксель по имени проформы и номеру таб. части. 
// 
// Параметры:
//  Name     - Строка 
//  Position - Число 
// 
// Возвращаемое значение:
//  Строка 
// 
Function SheetName(Name, Position) Export
	
	shName = Name + "_" + Position;
	
	Return shName;
	
EndFunction	// SheetName()

// Функция возвращает имя листа, содержащего источники выпадающих списков. 
// 
// Возвращаемое значение:
//  Строка.
// 
Function ListSourceSheetName() Export
	
	shName = "ИсточникСписков";
	
	return shName;
	
EndFunction // ListSourceSheetName()

// Функция выгружает на отдельный лист данные для установки выпадающих списков. 
// 
// Параметры:
//  Wb                - ComObject.
//  XlEnums           - Структура.
//  ДанныеДляВыгрузки - Структура.
// 
Function UnloadLists(Wb, XlEnums, ДанныеДляВыгрузки) Export
	
	LoadedLists = new Map;	
	
	// Создание и установка параметров для листа, являющегося источником выпадающих списков.
	ExcelSheet = Wb.Worksheets.Add(, Wb.Worksheets(Wb.Worksheets.Count));
	ExcelSheet.Name = ListSourceSheetName();
	ExcelSheet.StandardWidth           = 25;
	ExcelSheet.Cells.WrapText          = Истина;
	ExcelSheet.Cells.VerticalAlignment = XlEnums.xlConsts.xlCenter;
	ExcelSheet.Cells.NumberFormat      = "@";
	
	StartRow    = 2;
	StartColumn = 1;
    
	Сч = 0;
	For each kvp In ДанныеДляВыгрузки.СпискиВыгрузки Do
		
		CurColumn = StartColumn+Сч;			
		
		TypeStr = kvp.key;
        
		// Заголовок списка.
		Cell       = ExcelSheet.Cells(StartRow - 1, CurColumn);
		Cell.Value = TypeStr;
		
		// Выгрузка списка, со смещением на одну колонку выводятся ссылки на объекты в БД.
		List = kvp.Value;
		Счk = 0;
		For each item in List DO
			
			CurRow = StartRow+Счk;
			Cell = ExcelSheet.Cells(CurRow, CurColumn);
			Cell.Value = item.Presentation;
			Cell = ExcelSheet.Cells(CurRow, CurColumn + 1);
			Cell.Value = item.Value;
			
			Счk = Счk + 1;
            
		EndDo;	
		
		// Установка имени для области содержащей значения списка выбора для текущего типа.
		ListName = "List." + TypeStr;
		
		PosStart  = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow, CurColumn);
		PosEnd    = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, CurColumn);			
		Selection = ExcelSheet.Range(ExcelSheet.Cells(PosStart.Row, PosStart.Column),ExcelSheet.Cells(PosEnd.Row, PosEnd.Column));
		
		Selection.Name = ListName;
		
		LoadedLists.Insert(kvp.Key, ListName);
		
		Сч = Сч + 2;
		
	EndDo; 
	
	return LoadedLists;
	
EndFunction	// UnloadLists()

// Функция выполняет поиск списка, являющегося источником выпадающих списков. 
// 
// Параметры:
//  LoadedLists  - Соответствие.
//  MetaFields   - Структура.
//  PresentTypes - Массив.
// 
Function FindList(LoadedLists, MetaFields, PresentTypes) Export
	
	// Первые приоритет - поиск персонального списка по имени.
	ListName = LoadedLists[MetaFields.Имя];
	
	If ListName = Undefined Then
		
		// Второй приоритет - поиск списка по типу и виду представления.
		ListName = FindListOnCompositeKey(LoadedLists, MetaFields, PresentTypes);					
		
	EndIf;
	
	If ListName = Undefined Then
		
		// Третий приоритет - поиск списка по типу.
		ListName = LoadedLists[MetaFields.СтрТип];						
		
	EndIf;
	
	Return ListName;				
	
EndFunction // FindList()	

// Считывает скрытые параметры листа. 
// 
// Параметры:
//  ExcelSheet - ComObject.
// 
// ВозвращаемоеЗначение:
//  Структура.
// 
Function GetHiddenParams(ExcelSheet) Export
	
	res = New Structure;
	
	res.Insert("ИмяПроформы", String(ExcelSheet.Cells(1,1).Value));
	res.Insert("ВидЛиста"   , String(ExcelSheet.Cells(2,1).Value));
	res.Insert("ИмяТабЧасти", String(ExcelSheet.Cells(3,1).Value));
				
	return res;			
				
EndFunction // GetHiddenParams()

// Процедура выгружает в Эксель шаблоны проформ. 
// 
// Параметры:
//  ДанныеДляВыгрузки - Структура.
// 
Процедура ВыгрузитьШаблоныПроформ(ДанныеДляВыгрузки)  Экспорт
	
	Excel   = бит_ОбменДаннымиЭксельКлиентСервер.InitExcel(Истина);
	XlEnums = бит_ОбменДаннымиЭксельКлиентСервер.InitExcelEnums();
		
	Если НЕ Excel = Неопределено И ДанныеДляВыгрузки.Количество() > 0 Тогда
		
		// Параметры отображения
		ReprParams = new Structure;
		ReprParams.Insert("TableHeaderFontColor" , 16777215);  // Белый
		ReprParams.Insert("TableHeaderBackground", 16711680); // синий
		ReprParams.Insert("StandardWidth"        , 15);
		ReprParams.Insert("HeaderWidth"          , 25);
		ReprParams.Insert("TitleFontSize"        , 24);
		ReprParams.Insert("TableTitleFontSze"    , 16);
		
		Wb = Excel.Application.WorkBooks.Add(1);		
		
		// Загрузка данных для выпадающих списков
		LoadedLists = UnloadLists(Wb, XlEnums, ДанныеДляВыгрузки);
		
		// Создание шаблонов проформ
		UnloadTemplates(Wb, XlEnums, LoadedLists, ДанныеДляВыгрузки, ReprParams);
		
		// Лист "по-умолчанию" удаляем
		Wb.Worksheets(1).Delete();		
		
		Excel.Visible = Истина;
		
	КонецЕсли; 	
	
КонецПроцедуры // ВыгрузитьШаблонПроформы()

#КонецОбласти

#Область ВыгрузкаДанныхПроформИзExcel

// Функция конструирует имя листа эксель по имени проформы и номеру таб. части. 
// 
// Параметры:
//  Name     - Строка.
//  Position - Число.
// 
// Возвращаемое значение:
//  Строка.
// 
Function StandardFormSheetName(Name, TotalNumber, Number) Export
	
	shName = String(TotalNumber) + ". " + Name + "_" + Number;
	
	Return shName;
	
EndFunction	// StandardFormSheetName()

// Процедура выгружает в Эксель шаблоны проформ. 
// 
// Параметры:
//  ДанныеДляВыгрузки - Массив.
// 
Процедура ВыгрузитьПроформы(ДанныеДляВыгрузки)  Экспорт
	
	Excel   = бит_ОбменДаннымиЭксельКлиентСервер.InitExcel(Истина);
	XlEnums = бит_ОбменДаннымиЭксельКлиентСервер.InitExcelEnums();
	
	Если НЕ Excel = Неопределено И ДанныеДляВыгрузки.Количество() > 0 Тогда
		
		// Параметры отображения.
		ReprParams = new Structure;
		ReprParams.Insert("TableHeaderFontColor" , 16777215);  // Белый
		ReprParams.Insert("TableHeaderBackground", 26112);     // зеленый
		ReprParams.Insert("StandardWidth"        , 15);
		ReprParams.Insert("HeaderWidth"          , 25);
		ReprParams.Insert("TitleFontSize"        , 24);
		ReprParams.Insert("TableTitleFontSze"    , 16);
		NumberSplitChar = Excel.International(XlEnums.XlApplicationInternational.xlDecimalSeparator);		
		ReprParams.Insert("NumberSplitChar", NumberSplitChar);
		
		// Выгрузка проформ 
		UnloadStandardForms(Excel, XlEnums, ДанныеДляВыгрузки, ReprParams);
		
		Excel.Visible = Истина;
		
	КонецЕсли; 	
	
КонецПроцедуры // ВыгрузитьПроформы()

#КонецОбласти

#Область ЗагрузкаДанныхПроформИзExcel

// Функция конструктор структуры, моделирующей элемент списка-источника выпадающих списков.
// 
// Параметры:
//  Представление - Строка.
//  Ссылка        - Строка.
//  Значение      - Произвольный.
// 
// Возвращаемое значение:
//  Структура.
// 
Функция ЭлементСпискаИсточника(Представление, Ссылка, Значение = Неопределено)  Экспорт

	Элемент = Новый Структура("Представление, Ссылка, Значение"
	                            , Представление
								, Ссылка
								, Значение);

	Возврат Элемент;
	
КонецФункции // ЭлементСпискаИсточника()

// Функция конструктор структуры, моделирующей загруженной из эксель значение.
// 
// Параметры:
//  ЗначениеСтр  - Строка - Исходное значение.
//  МетаРеквизит - Структура.
//  ИмяСписка    - Строка.
//  Значение     - Произвольный - значение из БД.
// 
// Возвращаемое значение:
//  РезСтр - Структура.
// 
Функция ЗагруженноеЗначение(ЗначениеСтр, МетаРеквизит, ИмяСписка, Значение = Неопределено) Экспорт

	РезСтр = Новый Структура("ЗначениеСтр, Мета, ИмяСписка, Значение"
	                           , ЗначениеСтр
							   , МетаРеквизит
							   , ИмяСписка
							   , Значение);
	

	Возврат РезСтр;
	
КонецФункции // ЗагруженноеЗначение()

// Функция выполняет чтение списков, являющихся источниками выпадающих списков.
// 
// Параметры:
//  ЭксельКнига - ComObject.
// 
// Возвращаемое значение:
//  СпискиИсточники - Соответствие - Ключ:ИмяСписка; Значение:Список.
// 
Функция ПрочитатьСпискиИсточники(ЭксельКнига)  Экспорт
	
	СпискиИсточники = Новый Соответствие;
	
	Отказ = Ложь;
	
	ИмяЛистаИсточник = бит_ПроформыКлиентСервер.ListSourceSheetName();
	ЛистИсточник     = бит_ОбменДаннымиЭксельКлиентСервер.GetExcelSheet(ЭксельКнига, ИмяЛистаИсточник, Отказ);
	
	Если НЕ Отказ Тогда
		
		Arr = ЛистИсточник.UsedRange.Value;
		
		МассивКолонок = Arr.Unload();
		
		// Список-источник в Эксель состоит из двух колонок: 
		//  первая - представление, отображаемое в списке выбора;
		//  вторая - ссылка на соответствующий объект в 1С.
		// Поэтому считывать будет попарно.
		
		КоличествоКолонок = МассивКолонок.Количество();
		КоличествоСписков = КоличествоКолонок/2;
		
		Для н = 1 По КоличествоСписков Цикл
			
			// Выполняется чтение представлений и ссылок.
			Список = Новый Соответствие;
			
			МассивПредставления = МассивКолонок[(н-1)*2];
			МассивСсылки        = МассивКолонок[(н-1)*2+1];
			
			МаксИндексСсылки = МассивСсылки.ВГраница();
			ИмяСписка        = МассивПредставления[0];
			
			Для  Счк = 2 ПО МассивПредставления.Количество() Цикл
				
				Сч = Счк -1;
				ТекПредставление = МассивПредставления[Сч];
				
				Если ТекПредставление = Неопределено Тогда					
					Прервать;					
				КонецЕсли; 
				
				Если Сч <= МаксИндексСсылки Тогда
					
					ТекСсылка = МассивСсылки[Сч];
					
					Если ТекСсылка = Неопределено Тогда						
						ТекСсылка = "";						
					КонецЕсли; 
					
				Иначе	
					
					ТекСсылка = "";
					
				КонецЕсли; 
				
				// Упаковываем считанные данные в структуру.
				ЭлементСписка = бит_ПроформыКлиентСервер.ЭлементСпискаИсточника(ТекПредставление, ТекСсылка, Неопределено);
				Список.Вставить(ТекПредставление, ЭлементСписка);
				
			КонецЦикла; // По представлениям.
			
			СпискиИсточники.Вставить(ИмяСписка, Список);
			
		КонецЦикла; // По спискам.
		
	КонецЕсли; 
	
	// Для списка Булево сразу можно привести значения к типам 1С.
	СписокБулево = СпискиИсточники["Булево"];
	
	Если НЕ СписокБулево = Неопределено Тогда
		
		ТекЭлемент = СписокБулево["Да"];
		Если ТекЭлемент <> Неопределено Тогда			
			ТекЭлемент.Значение = Истина;			
        КонецЕсли; 
        
		ТекЭлемент = СписокБулево["Нет"];        
		Если ТекЭлемент <> Неопределено Тогда			
			ТекЭлемент.Значение = Ложь;			
		КонецЕсли; 
		
	КонецЕсли; 
	
	Возврат СпискиИсточники;
	
КонецФункции // ПрочитатьСписки()

// Процедура считывает из Эксель значения, соответствующие реквизитам шапки.
// Считывание производится по именам. Считанные данные записываются в структуру МодельДокумента.
// 
// Параметры:
//  Эксель           - ComObject.
//  ЭксельКнига      - ComObject.
//  ПсевдоМетаданные - Структура.
//  ИмяПроформы      - Строка.
//  МодельДокумента  - Структура - Ключ: ИмяРеквизита - Строка; Значение: Структура.ЗагруженноеЗначение.
// 
Процедура ПрочитатьДанныеШапки(ЭксельКнига, ЭксельЛистШапка, ПсевдоМетаданные, ИмяПроформы, МодельДокумента)  Экспорт
	
	Отказ = Ложь;
	
	Если ЭксельЛистШапка = Неопределено Тогда
		ИмяЛистаШапка = бит_ПроформыКлиентСервер.SheetName(ИмяПроформы, 1);
		ЛистШапка     = бит_ОбменДаннымиЭксельКлиентСервер.GetExcelSheet(ЭксельКнига, ИмяЛистаШапка, Отказ);
	Иначе	
		ЛистШапка = ЭксельЛистШапка;
	КонецЕсли; 
	
	Если НЕ Отказ Тогда
		
		ИменаЭксель = Новый Массив;
		
		Для каждого ИмяОбъект Из ЭксельКнига.Names Цикл
			
			ИменаЭксель.Добавить(ИмяОбъект.Name);
			
		КонецЦикла; 
		
		Для каждого МетаРеквизит Из ПсевдоМетаданные.Реквизиты Цикл
			
			Знч = ПрочитатьРеквизитШапки(ЛистШапка, ИменаЭксель, МетаРеквизит, ИмяПроформы);
			Если Знч <> Неопределено Тогда				
				МодельДокумента.Вставить(МетаРеквизит.Имя, Знч);				
			КонецЕсли; 
			
		КонецЦикла; 
		
		Для каждого МетаРеквизит Из ПсевдоМетаданные.СтандартныеРеквизиты Цикл
			
			Знч = ПрочитатьРеквизитШапки(ЛистШапка, ИменаЭксель, МетаРеквизит, ИмяПроформы);
			Если Знч <> Неопределено Тогда				
				МодельДокумента.Вставить(МетаРеквизит.Имя, Знч);				
			КонецЕсли; 
			
		КонецЦикла; 
		
	КонецЕсли; 	
	
КонецПроцедуры // ПрочитатьДанныеШапки()

// Процедура считывает данные табличных частей проформы из Excel.
// 
// Параметры:
//  ЭксельКнига       - ComObject.
//  ПсевдоМетаданные  - Структура.
//  ЗагруженныеДанные - Структура.
//  ИмяПроформы       - Строка.
// 
Процедура ПрочитатьДанныеТабЧастей(ЭксельКнига, ПсевдоМетаданные, ЗагруженныеДанные, ИмяПроформы)  Экспорт
	
	Отказ = Ложь;
	
	XlEnums	= бит_ОбменДаннымиЭксельКлиентСервер.InitExcelEnums();
	
	Для каждого ТекЛист Из ЭксельКнига.Sheets Цикл
		
		ПараметрыЛиста = бит_ПроформыКлиентСервер.GetHiddenParams(ТекЛист);
		
		Если ПараметрыЛиста.ИмяПроформы = ИмяПроформы Тогда
			
			Если ПсевдоМетаданные.ТабличныеЧасти.Свойство(ПараметрыЛиста.ИмяТабЧасти) Тогда
				
				МетаТабЧасть = ПсевдоМетаданные.ТабличныеЧасти[ПараметрыЛиста.ИмяТабЧасти];
				ДанныеТЧ = ПрочитатьДанныеТабЧасти(ТекЛист, МетаТабЧасть, ИмяПроформы);
				
				Если ДанныеТЧ.МассивДанных.Количество() > 0 Тогда
					
					ЗагруженныеДанные.Вставить(МетаТабЧасть.Имя, ДанныеТЧ);
					
				КонецЕсли; 
				
			КонецЕсли;
			
		КонецЕсли;  // Совпадает проформа 
		
	КонецЦикла; // ЛистыЭксель
	
КонецПроцедуры // ПрочитатьДанныеТабЧастей() 

// Функция данные табличной части проформы из Эксель.
// 
// Параметры:
//  ТекЛист      - ComObject.
//  МетаТабЧасть - Структура.
//  ИмяПроформы  - Строка.
// 
// Возвращаемое значение:
//  ДанныеТЧ - Структура.
// 
Функция ПрочитатьДанныеТабЧасти(ТекЛист,  МетаТабЧасть, ИмяПроформы)  Экспорт
	
	ДанныеТЧ = Новый Структура("Колонки, МассивДанных, ТаблицаДанных", Новый Соответствие, Новый Массив, Неопределено);
	
	Если ТипЗнч(ТекЛист) = Тип("ComObject") Тогда
		
		ИмяОблНачало = ИмяПроформы + "." + МетаТабЧасть.Имя;
		
		ЯчейкаНач = ТекЛист.Range(ИмяОблНачало);
		НомКол = ЯчейкаНач.Column;
		НомСтр  = ЯчейкаНач.Row;
		
		ЭтЗнч = ЯчейкаНач.Value;
		КоличествоКолонок = ТекЛист.Cells(1,1).SpecialCells(11).Column;
		КоличествоСтрок   = ТекЛист.Cells(1,1).SpecialCells(11).Row;
		
		
		РеквТЧ = Новый Соответствие;
		
		Для каждого МетаРекв Из МетаТабЧасть.Реквизиты Цикл
			
			РеквТЧ.Вставить(МетаРекв.Имя, МетаРекв);
			
		КонецЦикла; 
		
		НастройкиКолонок = Новый Соответствие;
		
		Для Счк = НомКол По КоличествоКолонок Цикл
			
			ИмяКолонки = ТекЛист.Cells(НомСтр, Счк).Value;
			
			Попытка
				
				ИмяСписка = ТекЛист.Cells(НомСтр + 1, Счк).Validation.Formula1;
				ИмяСписка = СтрЗаменить(ИмяСписка, "=List.", "");
				
			Исключение
				
				ИмяСписка = "";
				
			КонецПопытки;
			
			МетаРеквизит = РеквТЧ[ИмяКолонки];
			
			НастройкаКолонки = Новый Структура("Номер, Имя, ИмяСписка, Мета"
												, Счк-НомКол
												, ИмяКолонки
												, ИмяСписка
												, МетаРеквизит);
			
			// Сообщить(ИмяКолонки);
			НастройкиКолонок.Вставить(ИмяКолонки, НастройкаКолонки);
			
		КонецЦикла; 
		
		Selection = ТекЛист.Range(ТекЛист.Cells(НомСтр + 1, НомКол), ТекЛист.Cells(КоличествоСтрок, КоличествоКолонок));	
		Arr       = Selection.Value;
		
		МассивДанных = Arr.Unload();
		
		ДанныеТЧ.Колонки      = НастройкиКолонок;
		ДанныеТЧ.МассивДанных = МассивДанных;
		
	КонецЕсли; 	
	
	Возврат ДанныеТЧ;
	
КонецФункции // ПрочитатьДанныеТабЧасти()

#КонецОбласти

#Область ПрочиеПроцедурыИФункции

// Функция определяет, принадлежит ли реквизит, к числу жестко заданных реквизитов проформы.
// 
// Параметры:
//  ИмяРеквизита - Строка.
// 
// Возвращаемое значение:
//  Булево.
// 
Функция ЭтоФиксированныйРеквизит(ИмяРеквизита, ФиксированныеРеквизиты) Экспорт
	
	флЭтоФиксированный = ФиксированныеРеквизиты.НайтиПоЗначению(ВРег(ИмяРеквизита)) <> Неопределено;
	
	Возврат флЭтоФиксированный;
	
КонецФункции // ЭтоФиксированныйРеквизит()

// Функция дополняет код до необходимой длины слева.
// 
// Параметры:
//  ВхКод            - Строка.
//  Длина            - Строка.
//  СимволДополнение - Строка.
// 
// Возвращаемое значение:
//  Строка.
// 
Функция ДополнитьКод(ВхКод, Длина, СимволДополнение = "0")  Экспорт

	РезКод = ВхКод;
	
	Дельта = Длина - СтрДлина(ВхКод);
	
	Для Счн = 1 По Дельта Цикл
	
		РезКод = СимволДополнение + РезКод;
	
	КонецЦикла; 

	Возврат РезКод;
	
КонецФункции // ДополнитьКод()

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Процедура выводит заголовок листа. 
// 
// Параметры:
//  Sheet      - ComObject.
//  Title      - Строка.
//  FontSize   - Число.
//  Position   - Структура.
// 
Procedure SetSheetTitle(Sheet, Title, FontSize, Position)
	
	Cell = Sheet.Cells(Position.Row,Position.Column);
	Cell.Value = Title;
	Cell.Font.Size = FontSize;
	
EndProcedure // SetSheetTitle()	

// Процедура устанавливает стиль линий области реквизитов шапки. 
// 
// Параметры:
//  Sheet         - ComObject.
//  XlEnums       - Структура.
//  PositionStart - Структура.
//  PositionEnd   - Структура.
// 
Procedure SetHeaderLines(Sheet, XlEnums, PositionStart, PositionEnd)
	
	Selection = Sheet.Range(Sheet.Cells(PositionStart.Row, PositionStart.Column),Sheet.Cells(PositionEnd.Row, PositionEnd.Column));
	
	XlBordersIndex = XlEnums.XlBordersIndex;
	XlLineStyle    = XlEnums.XlLineStyle;
	XlBorderWeight = XlEnums.XlBorderWeight;
	XlNone         = XlEnums.XlConsts.xlNone;
	
	For each kvp in XlBordersIndex Do
		
		Border = Selection.Borders(kvp.value);
		
		If Find(kvp.key, "Edge") > 0 Then
			
			Border.LineStyle = XlLineStyle.xlContinuous;
			Border.Weight    = XlBorderWeight.xlMedium;
			
		EndIf;		
		
	EndDo;	
	
	For each kvp in XlBordersIndex Do
		
		Border = Selection.Borders(kvp.value);
		
		If Not Find(kvp.key, "Edge") > 0 Then
			
			Border.LineStyle = XlNone;
			
		EndIf;		
		
	EndDo;	
	
EndProcedure // SetHeaderLines() 	

// Процедура устанавливает стиль линий шапки таблицы. 
// 
// Параметры:
//  Sheet         - ComObject.
//  XlEnums       - Структура.
//  PositionStart - Структура.
//  PositionEnd   - Структура.
// 
Procedure SetTableHeaderLines(Sheet, XlEnums, PositionStart, PositionEnd)
	
	Selection = Sheet.Range(Sheet.Cells(PositionStart.Row, PositionStart.Column),Sheet.Cells(PositionEnd.Row, PositionEnd.Column));
	
	XlBordersIndex = XlEnums.XlBordersIndex;
	XlLineStyle    = XlEnums.XlLineStyle;
	XlBorderWeight = XlEnums.XlBorderWeight;
	XlNone         = XlEnums.XlConsts.xlNone;
	
	For each kvp in XlBordersIndex Do
		
		Border = Selection.Borders(kvp.value);
		
		If Find(kvp.key, "Edge") > 0 Then
			
			Border.LineStyle = XlLineStyle.xlContinuous;
			Border.Weight    = XlBorderWeight.xlThin;
			
		EndIf;			
		
	EndDo;	
	
EndProcedure // SetTableHeaderLines() 	

// Процедура устанавливает стиль ячейки шапки таблицы. 
// 
// Параметры:
//  Cell       - ComObject 
//  xlConst    - Структура 
//  Background - Число 
//  TextColor  - Число 
// 
Procedure SetTableHeaderStyle(Cell, xlConsts, Background, TextColor )
	
	Cell.VerticalAlignment   = xlConsts.xlCenter; 
	Cell.HorizontalAlignment = xlConsts.xlCenter;
	Cell.Interior.Color	     = Background;
	Cell.Font.Color          = TextColor;
	Cell.WrapText            = True;
	
EndProcedure // SetTableHeaderStyle()

// Процедура устанавливает стиль числовой ячейки. 
// 
// Параметры:
//  Cell    - ComObject.
//  xlConst - Структура.
// 
Procedure SetColumnNumberStyle(Cell, xlConsts)
	
	Cell.VerticalAlignment   = xlConsts.xlCenter; 
	Cell.HorizontalAlignment = xlConsts.xlCenter;
	Cell.WrapText = True;
	
EndProcedure // SetColumnNumberStyle()

// Процедура создает на листе Эксель поле для ввода реквизита шапки. 
// 
// Параметры:
//  ExcelSheet        - ComObject.
//  LoadedLists       - Соответствие.
//  XlEnums           - Структура.
//  МетаРеквизит      - Структура.
//  Pos               - Структура.
//  FormName          - Строка.
//  ВидыПредставлений - Массив.
// 
Procedure CreateHeaderField(Sheet, LoadedLists, XlEnums, МетаРеквизит, Pos, FormName, ВидыПредставлений)
	
	// Вывод синонимов полей.
	cell                   = Sheet.Cells(Pos.Row, Pos.Column);
	cell.Value             = МетаРеквизит.Синоним;
	cell.VerticalAlignment = XlEnums.xlConsts.xlCenter;
	
	// Установка имени области.
	cell      = Sheet.Cells(Pos.Row, Pos.Column + 1);
	cell.Name = FormName + "." + МетаРеквизит.Имя;
	
	// Установка списка выбора.
	ListName = FindList(LoadedLists, МетаРеквизит, ВидыПредставлений);	
	
	If ListName <> Undefined Then		
		бит_ОбменДаннымиЭксельКлиентСервер.SetList(cell, ListName, XlEnums);		
	EndIf;
	
EndProcedure // CreateHeaderField()	

// Функция выполняет поиск списка по составному ключу, 
// в котором учитывается тип поля и выражение представления. 
// 
// Параметры:
//  LoadedLists  - Соответствие.
//  MetaFields   - Структура.
//  PresentTypes - Массив.
// 
Function FindListOnCompositeKey(LoadedLists, MetaField, PresentTypes)
	
	ListName = Undefined;
	
	If ТипЗнч(MetaField.НастройкиОбмена) = Тип("Структура") Then
		
		StrFind   = MetaField.СтрТип + "/" + MetaField.НастройкиОбмена.ВыражениеПредставления;
		isPresent = PresentTypes.Find(StrFind);
		
		If isPresent <> Undefined Then
			
			Ключ     = MetaField.СтрТип + "_PresentationType_" + Формат((isPresent + 1), "ЧРГ=");
			ListName = LoadedLists[Ключ];			
			
		EndIf; 
		
	EndIf; 
	
	Return ListName;
	
EndFunction // FindListOnCompositeKey()	

// Процедура создает на листе Эксель область для ввода реквизитов шапки. 
// 
// Параметры:
//  ExcelSheet        - ComObject.
//  XlEnums           - Структура.
//  LoadedLists       - Соответствие.
//  ДанныеДляВыгрузки - Структура.
//  ReprParams        - Структура.
//  СтрПар            - Структура.
// 
Function CreateHeader(ExcelSheet ,XlEnums, LoadedLists, ДанныеДляВыгрузки, ReprParams, СтрПар)
	
	ПсевдоМета  = СтрПар.ПсевдоМетаданные;
	FormName    = ПсевдоМета.Имя;
	FormTitle   = СтрПар.Наименование;
	TablesCount = ПсевдоМета.ТабличныеЧасти.Количество();	
	ShNumber    = 1;
		
	StartRow    = 2;
	StartColumn = 3;
	
	// Заголовок.
	Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow, StartColumn);
	SetSheetTitle(ExcelSheet, FormTitle, 24, Position);
	
	// Реквизиты шапки.
	StartRow    = 4;
	StartColumn = 3;
    
	Сi = 0;
	
	// Создание полей для стандартных реквизитов.
	Для каждого МетаРеквизит Из ПсевдоМета.СтандартныеРеквизиты Цикл
		
		CurRow = StartRow + Сi;
		
		Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);				
		CreateHeaderField(ExcelSheet, LoadedLists, XlEnums, МетаРеквизит, Position, FormName, ДанныеДляВыгрузки.ВидыПредставлений);
		
		Сi = Сi + 1;
        
	КонецЦикла; // ПсевдоМета.Реквизиты.
	
	// Создание полей для реквизитов шапки.
	Для каждого МетаРеквизит Из ПсевдоМета.Реквизиты Цикл
		
		CurRow = StartRow + Сi;
		
		Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);				
		CreateHeaderField(ExcelSheet, LoadedLists, XlEnums, МетаРеквизит, Position, FormName, ДанныеДляВыгрузки.ВидыПредставлений);
		
		Сi = Сi + 1;
        
	КонецЦикла; // ПсевдоМета.Реквизиты.
	
	ExcelSheet.Columns(StartColumn).ColumnWidth   = ReprParams.HeaderWidth;
	ExcelSheet.Columns(StartColumn+1).ColumnWidth = ReprParams.HeaderWidth;
	
	// Установка стиля шапки.
	PositionStart = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow, StartColumn);
	PositionEnd   = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow  , StartColumn + 1);		
	SetHeaderLines(ExcelSheet, XlEnums, PositionStart, PositionEnd);
	
	LastPos = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);
	
	return LastPos;
	
EndFunction	// CreateHeader()

// Процедура создает шаблоны проформ в Эксель. 
// 
// Параметры:
//  ExcelSheet  - ComObject.
//  FormName    - Строка.
//  SheetType   - Строка.
//  TableName   - Строка.
//  RelisNumber - Строка.
// 
Procedure SetHiddenParams(ExcelSheet, FormName, SheetType, TableName, RelisNumber)
	
	ExcelSheet.Cells(1,1).Value       = FormName;
	ExcelSheet.Cells(2,1).Value       = SheetType;
	ExcelSheet.Cells(3,1).Value       = TableName;
	ExcelSheet.Cells(4,1).Value       = RelisNumber;
	ExcelSheet.Columns(1).ColumnWidth = 0;
	
EndProcedure // SetHiddenParams()	

// Процедура создает шаблоны проформ в Эксель. 
// 
// Параметры:
//  Wb                - ComObject.
//  XlEnums           - Структура.
//  LoadedLists       - Соответствие.
//  ДанныеДляВыгрузки - Структура.
//  ReprParams        - Структура.
// 
Procedure UnloadTemplates(Wb, XlEnums, LoadedLists, ДанныеДляВыгрузки, ReprParams)
	
		// Выгрузка шаблона проформ.
		Для каждого СтрПар Из ДанныеДляВыгрузки.МассивДанных Цикл
			
			ПсевдоМета  = СтрПар.ПсевдоМетаданные;
			FormName    = ПсевдоМета.Имя;
			FormTitle   = СтрПар.Наименование;
			TablesCount = ПсевдоМета.ТабличныеЧасти.Количество();
						
			ShNumber = 1;
			ShName   =  SheetName(FormName, ShNumber);
			
			// Лист для данных шапки и первой табличной части.
			ExcelSheet = Wb.Worksheets.Add(Wb.Worksheets(Wb.Worksheets.Count));
			ExcelSheet.Name          = ShName;
			ExcelSheet.StandardWidth = ReprParams.StandardWidth;
			
			// Вывод служебных данных.
			SetHiddenParams(ExcelSheet, СтрПар.ПсевдоМетаданные.Имя, "Шапка", "", ДанныеДляВыгрузки.НомерРелизаБИТ);
			
			LastPos = CreateHeader(ExcelSheet, XlEnums, LoadedLists, ДанныеДляВыгрузки, ReprParams, СтрПар);
			
			Сч = 1;
			Для каждого КиЗ Из СтрПар.ПсевдоМетаданные.ТабличныеЧасти Цикл
				
				МетаТаб = КиЗ.Значение;
				
				Если Сч > 1 Тогда
					
					// Создание листов для других табличных частей.
					ShNumber = ShNumber + 1;
					ShName = SheetName(FormName, ShNumber);
					
					ExcelSheet = Wb.Worksheets.Add(Wb.Worksheets(Wb.Worksheets.Count));
					ExcelSheet.Name          = ShName;
					ExcelSheet.StandardWidth = ReprParams.StandardWidth;
					
					Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(2, 3);
					SetSheetTitle(ExcelSheet, FormTitle, 24, Position);
					
				КонецЕсли; 
				
				// Создание шапок таблиц.
				If Сч = 1 Then
					
					StartRow = LastPos.Row + 3;
			        SetHiddenParams(ExcelSheet, СтрПар.ПсевдоМетаданные.Имя, "Шапка", МетаТаб.Имя, ДанныеДляВыгрузки.НомерРелизаБИТ);
					
				Else
					
					StartRow = 5;
					
					// Вывод служебных данных.
					SetHiddenParams(ExcelSheet, СтрПар.ПсевдоМетаданные.Имя, "ТабЧасть", МетаТаб.Имя, ДанныеДляВыгрузки.НомерРелизаБИТ);
					
				EndIf;
				
				StartColumn = 3;
				CurColumn = StartColumn;
				
				Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow - 1, StartColumn);
				SetSheetTitle(ExcelSheet, МетаТаб.Синоним, 16, Position);
				
				Счk = 1;
                ExcelSheet.Cells(StartRow + 2, CurColumn).Name = FormName + "." + МетаТаб.Имя;				
				Для каждого МетаРеквизит ИЗ МетаТаб.Реквизиты  Цикл
					
					Cell = ExcelSheet.Cells(StartRow, CurColumn);
					Cell.Value = МетаРеквизит.Синоним;
					SetTableHeaderStyle(Cell,XlEnums.xlConsts, ReprParams.TableHeaderBackground, ReprParams.TableHeaderFontColor);				
					
					Cell = ExcelSheet.Cells(StartRow + 1, CurColumn);
					Cell.Value = Счk;
					SetColumnNumberStyle(Cell, XlEnums.xlConsts);
					
					ExcelSheet.Cells(StartRow + 2, CurColumn).Value = МетаРеквизит.Имя;
					CurColumn = CurColumn + 1;
					
					// Установка списков
					ListName = FindList(LoadedLists, МетаРеквизит, ДанныеДляВыгрузки.ВидыПредставлений);					
					
					If ListName <> Undefined Then
						
						PosStart = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow + 3, CurColumn - 1);
						
	                    Selection = ExcelSheet.Range(ExcelSheet.Cells(PosStart.Row, PosStart.Column)
						                             ,ExcelSheet.Cells(PosStart.Row, PosStart.Column).End(XlEnums.XlDirection.xlDown));						
						
					    бит_ОбменДаннымиЭксельКлиентСервер.SetList(Selection, ListName, XlEnums);
										 
					EndIf;
					
					Счk = Счk + 1;
                    
				КонецЦикла; 
				
				// Скрываем ряд с именами.
				ExcelSheet.Rows(StartRow + 2).RowHeight = 0;
				
				// Установка сетки шапки таблицы.
				PosStart  = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow    , StartColumn);
				PosFinish = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow + 1, CurColumn - 1);			
				SetTableHeaderLines(ExcelSheet,XlEnums,PosStart, PosFinish);
				
				Сч = Сч + 1;
                
			КонецЦикла; 			
			
		КонецЦикла; // ДанныеДляВыгрузки	
	
EndProcedure // UnloadTemplates()	

// Процедура устанавливает значение и формат ячейки Excel. 
// 
Procedure SetCellValue(Cell, CurValue, TypeDescr, xlConsts, NumberSplitChar)
	
	If TypeOf(CurValue) = Type("String") Then
		
		Cell.Value = CurValue;
		
	ElsIf  TypeOf(CurValue) = Type("Date") Then
		
		Cell.Value = CurValue;
		
	ElsIf  TypeOf(CurValue) = Type("Number") Then
		
		Cell.Value = CurValue;
		
		If TypeDescr.ContainsType(Type("Number")) Then
			
			NumberDigits = TypeDescr.NumberQualifiers.FractionDigits;
			StrFormat    = "0";
			
			For Сi = 1 To NumberDigits Do
				
				If (Сi = 1) Then
					StrFormat = StrFormat + NumberSplitChar;
				EndIf;
				StrFormat = StrFormat + "0";
				
			EndDo;	
			
			Cell.NumberFormat  = StrFormat;
			
		EndIf;	
		
	Else
		
		Cell.Value = String(CurValue);
		
	EndIf;		 
	
	Cell.VerticalAlignment   = xlConsts.xlCenter; 
	Cell.HorizontalAlignment = xlConsts.xlGeneral;
	Cell.WrapText            = True;	
	
EndProcedure // SetCellValue()	

// Процедура создает на листе Эксель поле для ввода реквизита шапки. 
// 
// Параметры:
//  Sheet        - ComObject.
//  XlEnums      - Структура.
//  Pos          - Структура.
//  FormName     - Строка.
//  МетаРеквизит - Структура.
//  CurVal       - Произвольный.
// 
Procedure CreateStandardFormHeaderField(Sheet, XlEnums, Pos, FormName, МетаРеквизит, CurVal, NumberSplitChar)
	
	// Вывод синонимов полей
	cell = Sheet.Cells(Pos.Row, Pos.Column);
	cell.Value             = МетаРеквизит.Синоним;
	cell.VerticalAlignment = XlEnums.xlConsts.xlCenter;
	
	// Установка имени области
	cell = Sheet.Cells(Pos.Row, Pos.Column + 1);
	SetCellValue(cell, CurVal, МетаРеквизит.Тип, XlEnums.xlConsts, NumberSplitChar);
	cell.Name = FormName + "." + МетаРеквизит.Имя;
	
EndProcedure // CreateStandardFormHeaderField()	

// Процедура создает на листе Эксель область для ввода реквизитов шапки. 
// 
// Параметры:
//  ExcelSheet - ComObject.
//  XlEnums    - Структура.
//  ReprParams - Структура.
//  СтрПар     - Структура.
// 
Function CreateStandardFormHeader(ExcelSheet ,XlEnums, ReprParams, СтрПар)
	
	ПсевдоМета  = СтрПар.ПсевдоМетаданные;
	StFormData  = СтрПар.Данные;
	FormName    = ПсевдоМета.Имя;
	FormTitle   = СтрПар.Наименование;
	TablesCount = ПсевдоМета.ТабличныеЧасти.Количество();	
	ShNumber    = 1;	
	
	StartRow    = 2;
	StartColumn = 3;
	
	// Заголовок.
	Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow, StartColumn);
	SetSheetTitle(ExcelSheet, FormTitle, 24, Position);
	
	// Реквизиты шапки.
	StartRow    = 4;
	StartColumn = 3;		
	Сi          = 0;
	
	// Вывод стандартных реквизитов.
	Для каждого МетаРеквизит Из ПсевдоМета.СтандартныеРеквизиты Цикл
		
		CurRow = StartRow+Сi;
		
		Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);
		CurVal = Undefined;
		StFormData.Property(МетаРеквизит.Имя, CurVal);		
		CreateStandardFormHeaderField(ExcelSheet, XlEnums, Position, FormName, МетаРеквизит, CurVal, ReprParams.NumberSplitChar);
		
		Сi = Сi + 1;
        
	КонецЦикла; // ПсевдоМета.Реквизиты.
	
	// Вывод для реквизитов шапки.
	Для каждого МетаРеквизит Из ПсевдоМета.Реквизиты Цикл
		
		CurRow = StartRow + Сi;
		
		Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);	
		CurVal = Undefined;
		StFormData.Property(МетаРеквизит.Имя, CurVal);				
		CreateStandardFormHeaderField(ExcelSheet, XlEnums, Position, FormName, МетаРеквизит, CurVal, ReprParams.NumberSplitChar);
		
		Сi = Сi + 1;
        
	КонецЦикла; // ПсевдоМета.Реквизиты.
	
	ExcelSheet.Columns(StartColumn).ColumnWidth     = ReprParams.HeaderWidth;
	ExcelSheet.Columns(StartColumn + 1).ColumnWidth = ReprParams.HeaderWidth;
	
	// Установка стиля шапки.
	PositionStart = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow, StartColumn);
	PositionEnd   = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow  , StartColumn + 1);		
	SetHeaderLines(ExcelSheet, XlEnums, PositionStart, PositionEnd);
	
	LastPos = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow, StartColumn);
	
	return LastPos;
	
EndFunction	// CreateStandardFormHeader()

// Процедура создает шаблоны проформ в Эксель. 
// 
// Параметры:
//  Excel             - ComObject.
//  XlEnums           - Структура.
//  ДанныеДляВыгрузки - Массив.
//  ReprParams        - Структура.
// 
Procedure UnloadStandardForms(Excel, XlEnums, ДанныеДляВыгрузки, ReprParams)

	Wb = Excel.Application.WorkBooks.Add(1);	

	// Выгрузка шаблона проформ.
	ShTotalNumber = 0;		
	Для каждого СтрПар Из ДанныеДляВыгрузки Цикл
		
		ПсевдоМета  = СтрПар.ПсевдоМетаданные;
		Данные      = СтрПар.Данные;
		FormName    = ПсевдоМета.Имя;
		FormTitle   = СтрПар.Наименование;
		TablesCount = ПсевдоМета.ТабличныеЧасти.Количество();
		
		ShNumber      = 1;
        ShTotalNumber = ShTotalNumber + 1;					
		ShName        = StandardFormSheetName(FormName, ShTotalNumber, ShNumber);
		
		// Лист для данных шапки и первой табличной части.
		ExcelSheet = Wb.Worksheets.Add(Wb.Worksheets(Wb.Worksheets.Count));
		ExcelSheet.Name          = ShName;
		ExcelSheet.StandardWidth = ReprParams.StandardWidth;
		
		// Вывод служебных данных.
		SetHiddenParams(ExcelSheet,СтрПар.ПсевдоМетаданные.Имя, "Шапка", "", СтрПар.НомерРелизаБИТ);
        
        // Вывод данных шапки.
		LastPos = CreateStandardFormHeader(ExcelSheet, XlEnums, ReprParams, СтрПар);
		
		// Вывод данных табличных частей.
		Сч = 1;
		Для каждого КиЗ Из СтрПар.ПсевдоМетаданные.ТабличныеЧасти Цикл
			
			МетаТаб = КиЗ.Значение;
			
			Если Сч > 1 Тогда
				
				// Создание листов для других табличных частей.
				ShNumber      = ShNumber + 1;
                ShTotalNumber = ShTotalNumber + 1;					
		        ShName        = StandardFormSheetName(FormName, ShTotalNumber, ShNumber);
				
				ExcelSheet = Wb.Worksheets.Add(Wb.Worksheets(Wb.Worksheets.Count));
				ExcelSheet.Name          = ShName;
				ExcelSheet.StandardWidth = ReprParams.StandardWidth;
				
				Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(2, 3);
				SetSheetTitle(ExcelSheet, FormTitle, 24, Position);
				
			КонецЕсли; 
			
			// Создание шапок таблиц.
			If Сч = 1 Then
				
				StartRow = LastPos.Row + 3;
				SetHiddenParams(ExcelSheet,СтрПар.ПсевдоМетаданные.Имя, "Шапка", МетаТаб.Имя, СтрПар.НомерРелизаБИТ);
				
			Else
				
				StartRow = 5;
				
				// Вывод служебных данных.
				SetHiddenParams(ExcelSheet,СтрПар.ПсевдоМетаданные.Имя, "ТабЧасть", МетаТаб.Имя, СтрПар.НомерРелизаБИТ);
				
			EndIf;
			
			StartColumn = 3;
			CurColumn   = StartColumn;
			
			Position = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow - 1, StartColumn);
			SetSheetTitle(ExcelSheet, МетаТаб.Синоним, 16, Position);
			
			Счk = 1;
			ExcelSheet.Cells(StartRow + 2, CurColumn).Name = FormName + "." + МетаТаб.Имя;				
			Для каждого МетаРеквизит ИЗ МетаТаб.Реквизиты  Цикл
				
				Cell = ExcelSheet.Cells(StartRow,CurColumn);
				Cell.Value = МетаРеквизит.Синоним;
				SetTableHeaderStyle(Cell, XlEnums.xlConsts, ReprParams.TableHeaderBackground, ReprParams.TableHeaderFontColor);				
				
				Cell = ExcelSheet.Cells(StartRow + 1, CurColumn);
				Cell.Value = Счk;
				SetColumnNumberStyle(Cell, XlEnums.xlConsts);
				
				ExcelSheet.Cells(StartRow + 2, CurColumn).Value = МетаРеквизит.Имя;
				CurColumn = CurColumn + 1;
				
				Счk = Счk + 1;
                
			КонецЦикла; 
			
			// Скрываем ряд с именами.
			ExcelSheet.Rows(StartRow + 2).RowHeight = 0;
			
			// Установка сетки шапки таблицы.
			PosStart  = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow    , StartColumn);
			PosFinish = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow + 1, CurColumn - 1);			
			SetTableHeaderLines(ExcelSheet, XlEnums, PosStart, PosFinish);
			
			// Вывод данных таблиц.
			ИмяТаблицы = МетаТаб.Имя;
			
			Если Данные.Свойство(ИмяТаблицы) Тогда
				
				 StartRow = StartRow + 3;
				 CurRow   = StartRow;
                 
                 // Закрепление шапки таблицы.
				 Selection = ExcelSheet.Range(ExcelSheet.Cells(StartRow, StartColumn - 1), ExcelSheet.Cells(StartRow, StartColumn - 1)).Select();
				 Excel.Application.ActiveWindow.FreezePanes = True;
				 
				 // Вывод строк таблицы.
				 Для каждого МодельСтроки Из Данные[ИмяТаблицы] Цикл
					 
				     CurColumn = StartColumn;						 
					 Для каждого МетаРеквизит ИЗ МетаТаб.Реквизиты  Цикл
						 
						 // Получение текущего значения.
						 ТекЗнч = Неопределено;
						 МодельСтроки.Свойство(МетаРеквизит.Имя, ТекЗнч);
						 Если ТекЗнч = Неопределено Тогда						 
						 	ТекЗнч = "";						 
						 КонецЕсли; 
						 
						 // Установка значения ячейки и форматирование.
                         Cell = ExcelSheet.Cells(CurRow, CurColumn);							 
						 SetCellValue(Cell, ТекЗнч, МетаРеквизит.Тип, XlEnums.xlConsts, ReprParams.NumberSplitChar); 
						 CurColumn = CurColumn + 1;
						 
					 КонецЦикла; // По колонкам таб. части.
					 
					 CurRow = CurRow + 1;
                     
				 КонецЦикла; // По строкам таб. части.
				 
				 // Установка сетки таблицы.
				 PosStart  = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(StartRow  , StartColumn);
				 PosFinish = бит_ОбменДаннымиЭксельКлиентСервер.CellPosition(CurRow - 1, CurColumn - 1);			
				 SetTableHeaderLines(ExcelSheet, XlEnums, PosStart, PosFinish);
				 
			КонецЕсли; // ТабЧастьЕсть.
			
			Сч = Сч + 1;
            
		КонецЦикла; // По табличным частям.			
		
	КонецЦикла; // ДанныеДляВыгрузки.
	
	// Лист "по-умолчанию" удаляем.
	Wb.Worksheets(Wb.Worksheets.Count).Delete();				
		
EndProcedure // UnloadStandardForms()	

// Функция считывает из Эксель значение, соответствующее реквизиту шапки документа.
// Считывание производится по именам.
// 
// Параметры:
//  Лист         - ComObject.
//  ИменаЭксель  - Массив.
//  МетаРеквизит - Структура.
//  ИмяПроформы  - Строка.
// 
Функция ПрочитатьРеквизитШапки(Лист, ИменаЭксель, МетаРеквизит, ИмяПроформы)
	
	Знч       = Неопределено;
	ИмяЭксель = ИмяПроформы + "." + МетаРеквизит.Имя;
	
	Если ИменаЭксель.Найти(ИмяЭксель) <> Неопределено Тогда
		
		// Получение значения по имени.
		ЗнчСтр    = Лист.Range(ИмяЭксель).Value;
		
		Попытка
			
			// Определим, установлен ли для данной ячейки список выбора, если установлен, то какой.
			ИмяСписка = Лист.Range(ИмяЭксель).Validation.Formula1;
			ИмяСписка = СтрЗаменить(ИмяСписка, "=List.", "");
			
		Исключение
			
			ИмяСписка = "";
			
		КонецПопытки;
		
		Знч = бит_ПроформыКлиентСервер.ЗагруженноеЗначение(ЗнчСтр, МетаРеквизит, ИмяСписка);
		
	КонецЕсли; 
	
	Возврат Знч;
	
КонецФункции // ПрочитатьРеквизитШапки() 

#КонецОбласти
