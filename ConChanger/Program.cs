// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;


Console.WriteLine("Welcome to the ConConverter");

// The Con workbook beeing read
var wb = new XLWorkbook("C:\\Users\\tguiv\\OneDrive\\Desktop\\New folder\\DAS 18639.XLSX");

int ConLineCount = CountLines(wb);

// the new workbook with the data in it
var wbout = new XLWorkbook();
wbout.AddWorksheet();



var value = wb.Worksheet(2).Cell(4, 4).Value;
wbout.Worksheet(1).Cell(1, 1).Value = value;
wbout.SaveAs("C:\\Users\\tguiv\\OneDrive\\Desktop\\New folder\\out.XLSX");



static int CountLines (XLWorkbook wb)
    {
    string ContSheetName = "Continuation Sheet";
    IXLWorksheet ContSheet = wb.Worksheet(3);
   if (ContSheet.Name != ContSheetName) { throw new ArgumentException("The continuation sheet wasnt found"); }

    for (int i = 4;i<100; i++)
    {
        if (ContSheet.Cell(i,2).Value.IsBlank)
        { return i - 4; }
    }

    return 0;

}