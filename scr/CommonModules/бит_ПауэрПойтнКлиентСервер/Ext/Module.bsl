
#Область СлужебныйПрограммныйИнтерфейс

// Функция создает объект PowerPoint.
// 
// Параметры:
//  DisplayAlerts - Булево.
//  Visible       - Булево.
//  Fail          - Булево
// 
// Возвращаемое значение:
//  PowerPoint - ComObject.
// 
Function InitPowerPoint(DisplayAlerts = False, Visible = True, Fail = False) Export
	
	PowerPointApp = Undefined;
	
	Try
		
		// Создание объекта Microsoft PowerPoint.
		PowerPointApp = new COMОбъект("PowerPoint.Application");
		PowerPointApp.DisplayAlerts = DisplayAlerts;
		PowerPointApp.Visible = Visible;
	
	Except
		
		Fail = True;
		
		ТекстСообщения = НСтр("ru = 'Не удалось создать объект PowerPoint по причине: %1%.'");
		ТекстСообщения = бит_ОбщегоНазначенияКлиентСервер.ПодставитьПараметрыСтроки(ТекстСообщения, Строка(ОписаниеОшибки()));
		бит_ОбщегоНазначенияКлиентСервер.ВывестиСообщение(ТекстСообщения);
		
	EndTry;
	
	Return PowerPointApp;
	
EndFunction // InitPowerPoint()

// Процедура закрывает PowerPoint.
// 
// Параметры:
//  PowerPointApp - ComObject.
// 
Procedure QuitPowerPoint(PowerPointApp) Export
	
	If NOT PowerPointApp = Undefined Then
		
		PowerPointApp.Quit();
		PowerPointApp = Undefined;
		
	EndIf;
	
EndProcedure  // QuitPowerPoint() 

// Функция открывает PowerPoint файл.
// 
// Параметры:
//  PowerPoint - ComObject("PowerPoint.Application").
//  Path       - Строка, полный путь где лежит файл.
//  Fail 	   - Булево, по умолчанию Ложь.
// 
// Возвращаемое значение:
//   PowerPointFile - ComObject, в случае неудачи Неопределено.
// 
Function OpenPowerPointFile(PowerPointApp, Path, Fail = False) Export
	
	PowerPointFile = Undefined;		
	
	If NOT PowerPointApp = Undefined Then
		
		Try
			
			PowerPointFile = PowerPointApp.Presentations.Open(Path);
			
		Except
			
			Fail = True;
			
			ТекстСообщения = НСтр("ru = 'Не удалось открыть PowerPoint файл ""%1%"" по причине: %2%'");
			ТекстСообщения = бит_ОбщегоНазначенияКлиентСервер.ПодставитьПараметрыСтроки(ТекстСообщения, Path, Строка(ОписаниеОшибки()));
			бит_ОбщегоНазначенияКлиентСервер.ВывестиСообщение(ТекстСообщения);
			
		EndTry;
		
	EndIf;
	
	Return PowerPointFile;
	
EndFunction // ОткрытьPowerPointФайл()

// Процедура закрывает книгу эксель.
// 
// Параметры:
//  ExcelWorkbook - ComObject.
//  SaveChanges   - Булево, по умолчанию Ложь.
// 
Procedure ClosePowerPointFile(PowerPointWorkbook, SaveChanges) Export
	
	If NOT PowerPointWorkbook = Undefined Then
		
		PowerPointWorkbook.Close(SaveChanges);
		PowerPointWorkbook = Undefined;
		
	EndIf;
	
EndProcedure

#КонецОбласти
