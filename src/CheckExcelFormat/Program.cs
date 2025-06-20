using System.Diagnostics;
using ClosedXML.Excel;

ReadViaClosedXml();

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
