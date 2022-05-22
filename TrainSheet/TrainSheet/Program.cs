using DevExpress.Xpo;
using DevExpress.Xpo.DB;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using TrainSheet;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

//Создание связи с базой данных MS SQL Server.
string sqlConn =
   MSSqlConnectionProvider.GetConnectionString(@"localhost\SQLEXPRESS", "WheelSheet");
XpoDefault.DataLayer = XpoDefault.GetDataLayer(sqlConn, AutoCreateOption.DatabaseAndSchema);

//Создание новой сессии.
Session session =  new Session();

//Базовое хранилище данных для запросов LINQ
XPQuery<WagonSheet> wagonSheets = new XPQuery<WagonSheet>(session);
XPQuery<WagonList> wagonLists = new XPQuery<WagonList>(session);

HELP:
Console.WriteLine("Здравстуйте.\nКоманды для работы с программой:\n -list - Список доступных номеров поездов,\n -exit - Выход из программы.");

string str;
ERROR1:
Console.WriteLine("Введите команду или номер поезда. (Список команд: -help)");
switch (str = Console.ReadLine().Trim())
{
    case "-list":
        foreach (var w in wagonSheets)
        {
            Console.WriteLine(w.TrainNumber);
        }

        goto ERROR1;
            break;
    case "-exit":
        Environment.Exit(0);
            break;
    case "-help":
        goto HELP;
        break;
}

int trainNumber = 0;
if (int.TryParse(str, out trainNumber))
{
    trainNumber = int.Parse(str);
}
else
{
    Console.WriteLine("Некоррекный ввод");
    goto ERROR1;
}

bool result = wagonSheets.Any(w => w.TrainNumber == trainNumber);
if (!result)
{
    Console.WriteLine("Данный номер поезда не существует.");
    goto ERROR1;
}

Console.WriteLine("Идёт формирование отчёта.");

//Создание объекта приложения Excel.
Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

//Количество листов в рабочей книге.
app.SheetsInNewWorkbook = 1;

//Добавить рабочую книгу.
Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

//Отключение отображения окон с сообщениями.
app.DisplayAlerts = false;

//Получение листа документа и наименование его.
Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.get_Item(1);
worksheet.Name = "Натурный лист поезда";

//Устанавливание формата ячеек для листа.
ExcelRange range = worksheet.Cells;
range.Cells.Font.Name = "Times New Roman";
range.Cells.Font.Size = 10;
range.HorizontalAlignment = Constants.xlCenter;

Console.WriteLine("Формирование шапки отчёта.");
//Заголовок натурного листа.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[1, 1], (ExcelRange)worksheet.Cells[1, 7]);
range.Merge(Type.Missing);
range.Cells.Font.Name = "Times New Roman";
range.Cells.Font.Size = 18;
range.Cells.Font.Bold = true;
range.HorizontalAlignment = Constants.xlCenter;
range.VerticalAlignment = Constants.xlTop;
range.Value2 = "НАТУРНЫЙ ЛИСТ ПОЕЗДА";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[3, 1], (ExcelRange)worksheet.Cells[3, 1]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = "Поезд №:";

//Внесение номера поезда в отчёт.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[3, 3], (ExcelRange)worksheet.Cells[3, 3]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = trainNumber;

range = worksheet.get_Range((ExcelRange)worksheet.Cells[3, 4], (ExcelRange)worksheet.Cells[3, 4]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = "Станция:";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[4, 1], (ExcelRange)worksheet.Cells[4, 1]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = "Состав №:";

//Формат шапки натурного листа.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 1], (ExcelRange)worksheet.Cells[6, 7]);
range.Cells.Font.Bold = true;
range.HorizontalAlignment = Constants.xlCenter;
range.VerticalAlignment = Constants.xlCenter;
range.WrapText = true;

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 1], (ExcelRange)worksheet.Cells[6, 1]);
range.Value2 = "№";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 2], (ExcelRange)worksheet.Cells[6, 2]);
range.Value2 = "№ вагона";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 3], (ExcelRange)worksheet.Cells[6, 3]);
range.Value2 = "Накладная";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 4], (ExcelRange)worksheet.Cells[6, 4]);
range.Value2 = "Дата операции";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 5], (ExcelRange)worksheet.Cells[6, 5]);
range.Value2 = "Груз";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 6], (ExcelRange)worksheet.Cells[6, 6]);
range.Orientation = 90;
range.Value2 = "Вес по документам (т)";

range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 7], (ExcelRange)worksheet.Cells[6, 7]);
range.Value2 = "Последняя операция";

//Получение элемента WagonSheet с номером, указанным пользователем.
WagonSheet colletion1 = wagonSheets.First(wagonSheets => wagonSheets.TrainNumber == trainNumber);

//Вносим номер состава в отчёт.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[4, 3], (ExcelRange)worksheet.Cells[4, 3]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = colletion1.WagonNumber;

//Вносим название станции дислокации.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[3, 5], (ExcelRange)worksheet.Cells[3, 5]);
range.HorizontalAlignment = Constants.xlLeft;
range.Value2 = colletion1.LastStationName;

//Сортируем по позиции вагона в поезде.
var colletion = colletion1.WagonLists.OrderBy(item => item.PositionInTrain);

Console.Write("Заполнение отчёта данными.");
//Вносим данные по вагонам в отчёт.
int row = 7, col = 1;
foreach (WagonList item in colletion)
{
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 1], (ExcelRange)worksheet.Cells[row, 1]);
    range.Value2 = col++;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 2], (ExcelRange)worksheet.Cells[row, 2]);
    range.Value2 = item.CarNumber;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 3], (ExcelRange)worksheet.Cells[row, 3]);
    range.Value2 = item.InvoiceNum;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 4], (ExcelRange)worksheet.Cells[row, 4]);
    range.Value2 = item.WhenLastOperation;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 5], (ExcelRange)worksheet.Cells[row, 5]);
    range.Value2 = item.FreightEtsngName;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 6], (ExcelRange)worksheet.Cells[row, 6]);
    range.Value2 = item.FreightTotalWeightKg / 1000;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 7], (ExcelRange)worksheet.Cells[row, 7]);
    range.Value2 = item.LastOperationName;
    row++;
    if (row % 5 == 0)
        Console.Write(".");
}
Console.WriteLine();

// Изменение формата для значений веса и даты.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[7, 6], (ExcelRange)worksheet.Cells[row - 1, 6]);
range.NumberFormat = "#.00";
range = worksheet.get_Range((ExcelRange)worksheet.Cells[7, 4], (ExcelRange)worksheet.Cells[row - 1, 4]);
range.NumberFormat = "m/d/yyyy";

//Запрос для расчётов в конце вывода списка. 
var sgt = from c in colletion
          group c by c.FreightEtsngName into cc
          select new { Name = cc.Key, Count = cc.Count(), TotalWeight = cc.Sum(item => item.FreightTotalWeightKg)};

//Внесение расчётов в конец вывода списка.
foreach (var s in sgt)
{
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 2], (ExcelRange)worksheet.Cells[row, 2]);
    range.Value2 = s.Count;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 5], (ExcelRange)worksheet.Cells[row, 5]);
    range.Value2 = s.Name;
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 6], (ExcelRange)worksheet.Cells[row, 6]);
    range.Value2 = s.TotalWeight / 1000;
    row++;
}

//Конечные рассчёты.
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 1], (ExcelRange)worksheet.Cells[row, 2]);
    range.Merge();
    range.Value2 = "Всего: " + sgt.Sum(item => item.Count).ToString();
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 5], (ExcelRange)worksheet.Cells[row, 5]);
    range.Value2 = sgt.Count();
    range = worksheet.get_Range((ExcelRange)worksheet.Cells[row, 6], (ExcelRange)worksheet.Cells[row, 6]);
    range.Value2 = sgt.Sum(item => item.TotalWeight) / 1000;

//Установление формата шрифта для конечных расчётов.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[row - sgt.Count(), 1], (ExcelRange)worksheet.Cells[row, 7]);
range.Cells.Font.Bold = true;

//Установка рамок таблицы.
range = worksheet.get_Range((ExcelRange)worksheet.Cells[6, 1], (ExcelRange)worksheet.Cells[row, 7]);
range.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous;
range.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous;
range.Borders.get_Item(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous;
range.Borders.get_Item(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlContinuous;
range.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
range.Borders.Color = Color.Black;

//Изменяет ширину столбцов в диапазоне или высоту строк в диапазоне, чтобы достичь наилучшей посадки.
range = worksheet.get_Range(worksheet.Columns, worksheet.Columns);
range.AutoFit();
range = worksheet.get_Range(worksheet.Rows, worksheet.Rows);
range.AutoFit();

Console.WriteLine("Готово.");
//Отобразить Excel
app.Visible = true;

Console.WriteLine();
Console.WriteLine();
Console.WriteLine();

goto HELP;