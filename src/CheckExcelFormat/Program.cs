using System.Diagnostics;
using ClosedXML.Excel;
using OfficeOpenXml;

ReadViaClosedXml();
ReadViaEPPlus();

static void ReadViaClosedXml()
{
    Console.WriteLine("ReadViaClosedXml");
    var stopWatch = new Stopwatch();
    stopWatch.Start();

    var workBook = new XLWorkbook("sample.xlsx");
    var workSheet = workBook.Worksheet("依頼登録");
    var firstRow = 6;
    var maxRow = 40;
    var firstColumn = 1;
    var maxColumn = 50;

    for (var row = firstRow; row < maxRow; row++)
    {
        for (var column = firstColumn; column < maxColumn; column++)
        {
            var cell = workSheet.Cell(row, column);
            Console.WriteLine($"row: {row}, column: {column}, value: {cell.Value}");
        }
    }

    stopWatch.Stop();
    Console.WriteLine($"Elapsed: {stopWatch.Elapsed}");
}

static void ReadViaEPPlus()
{
    Console.WriteLine("ReadViaEPPlus");

    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    var firstRow = 6;
    var maxRow = 40_000;
    var firstColumn = 1;
    var maxColumn = 52;

    (var valueTime, var value) = ReadValue(firstRow, maxRow, firstColumn, maxColumn);
    (var textTime, var text) = ReadText(firstRow, maxRow, firstColumn, maxColumn);
    Console.WriteLine($"Value: {valueTime}, last value: {value}");
    Console.WriteLine($"Text: {textTime}, last text: {text}");
}

static (TimeSpan elapsed, string? lastValue) ReadValue(int firstRow, int maxRow, int firstColumn, int maxColumn)
{
    using var package = new ExcelPackage(new FileInfo("sample.xlsx"));
    using var workSheet = package.Workbook.Worksheets["依頼登録"];
    var stopWatch = new Stopwatch();
    stopWatch.Start();

    string? value = null;
    for (var row = firstRow; row < maxRow; row++)
    {
        for (var column = firstColumn; column < maxColumn; column++)
        {
            var cell = workSheet.Cells[row, column];
            value = cell.Value?.ToString();
        }
    }

    stopWatch.Stop();
    return (stopWatch.Elapsed, value);
}

static (TimeSpan elapsed, string? lastText) ReadText(int firstRow, int maxRow, int firstColumn, int maxColumn)
{
    using var package = new ExcelPackage(new FileInfo("sample.xlsx"));
    using var workSheet = package.Workbook.Worksheets["依頼登録"];
    var stopWatch = new Stopwatch();
    stopWatch.Start();

    string? text = null;
    for (var row = firstRow; row < maxRow; row++)
    {
        for (var column = firstColumn; column < maxColumn; column++)
        {
            var cell = workSheet.Cells[row, column];
            text = cell.Text;
        }
    }

    stopWatch.Stop();
    return (stopWatch.Elapsed, text);
}
