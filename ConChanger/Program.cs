// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;


Console.WriteLine("Welcome to the ConConverter");

// The Con workbook beeing read
XLWorkbook wb = new XLWorkbook("C:\\Users\\tguiv\\OneDrive\\Desktop\\New folder\\DAS 18639.XLSX");
int ConLineCount = CountLines(wb);

Console.WriteLine("Opened the Concession and found " + ConLineCount.ToString() + " lines" );


// the new workbook to put the rearanged data into
var wbout = new XLWorkbook();
wbout.AddWorksheet();

Headers(wbout);
ConSheet(wb, wbout, ConLineCount);
ContinuationSheet(wb, wbout, ConLineCount);


//saving the new workbook
wbout.SaveAs("C:\\Users\\tguiv\\OneDrive\\Desktop\\New folder\\out.XLSX");

static void ContinuationSheet(XLWorkbook wb, XLWorkbook wbout, int ConLineCount)
{
    IXLWorksheet WSout = wbout.Worksheets.First();
    IXLWorksheet ContSheet = wb.Worksheet(3);
    string ContSheetName = "Continuation Sheet";
    if (ContSheet.Name != ContSheetName) { throw new ArgumentException("The continuation sheet wasnt found"); }

    ContSheet.Range(4, 2, ConLineCount+3, 2).CopyTo(WSout.Cell(2, 20)); // Item No.
    ContSheet.Range(4, 3, ConLineCount + 3, 3).CopyTo(WSout.Cell(2, 21));// Defect Location Code Group
    ContSheet.Range(4, 4, ConLineCount + 3, 4).CopyTo(WSout.Cell(2, 22)); // Defect Location Code Description
    ContSheet.Range(4, 5, ConLineCount + 3, 5).CopyTo(WSout.Cell(2, 23)); // Defect Type Code Group
    ContSheet.Range(4, 6, ConLineCount + 3, 6).CopyTo(WSout.Cell(2, 24)); //"Defect Type Code Description"
    ContSheet.Range(4, 7, ConLineCount + 3, 7).CopyTo(WSout.Cell(2, 25)); //"Cause Type Code Group"
    ContSheet.Range(4, 8, ConLineCount + 3, 8).CopyTo(WSout.Cell(2, 26)); //"Casue Type Code Description"
    ContSheet.Range(4, 9, ConLineCount + 3, 9).CopyTo(WSout.Cell(2, 27)); //"Zone"
    ContSheet.Range(4, 10, ConLineCount + 3, 10).CopyTo(WSout.Cell(2, 28)); //"Sheet"
    ContSheet.Range(4, 11, ConLineCount + 3, 11).CopyTo(WSout.Cell(2, 29)); //"Nominal"
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 30)); // "Toll(-)"
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 31)); // Toll(+)
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 32)); // Actual
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 33));
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 33));
    ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 33));
}

static void ConSheet (XLWorkbook wb, XLWorkbook wbout, int ConLineCount)
{
    IXLWorksheet WSout = wbout.Worksheets.First();
    IXLWorksheet ConSheet = wb.Worksheet(2);
    string ConSheetName = "Concession or Deviation Permit";
    if (ConSheet.Name != ConSheetName) { throw new ArgumentException("The Con or DP sheet wasnt found"); }

    WSout.Range(2, 1,  ConLineCount,1).Value = ConSheet.Cell(4,4).Value; //1.1
    WSout.Range(2, 2, ConLineCount, 2).Value = ConSheet.Cell(6, 3).Value; //1.2
    WSout.Range(2, 3, ConLineCount, 3).Value = ConSheet.Cell(6, 8).Value; //1.2
    WSout.Range(2, 4, ConLineCount, 4).Value = ConSheet.Cell(4, 19).Value; //1.3
    WSout.Range(2, 5, ConLineCount, 5).Value = ConSheet.Cell(6, 20).Value; //1.4
    WSout.Range(2, 6, ConLineCount, 6).Value = ConSheet.Cell(9, 1).Value; //2.1
    WSout.Range(2, 7, ConLineCount, 7).Value = ConSheet.Cell(9, 6).Value; //2.2

}

static void Headers (XLWorkbook wbout)
{
    IXLWorksheet WSout = wbout.Worksheets.First();

    WSout.Cell(1, 1).Value = "1.1";
    WSout.Cell(1, 2).Value = "1.2";
    WSout.Cell(1, 3).Value = "1.2";
    WSout.Cell(1, 4).Value = "1.3";
    WSout.Cell(1, 5).Value = "1.4";
    WSout.Cell(1, 6).Value = "2.1";
    WSout.Cell(1, 7).Value = "2.2";
    WSout.Cell(1, 8).Value = "2.3";
    WSout.Cell(1, 9).Value = "2.4";
    WSout.Cell(1, 10).Value = "2.5";
    WSout.Cell(1, 11).Value = "2.6";
    WSout.Cell(1, 12).Value = "2.7";
    WSout.Cell(1, 13).Value = "2.8";
    WSout.Cell(1, 14).Value = "2.9";
    WSout.Cell(1, 15).Value = "Previous Subs of a similar";
    WSout.Cell(1, 16).Value = "Previous Subs to actual";
    WSout.Cell(1, 17).Value = "4";
    WSout.Cell(1, 17).Value = "Originator";

    WSout.Cell(1, 20).Value = "Item No.";
    WSout.Cell(1, 21).Value = "Defect Location Code Group";
    WSout.Cell(1, 22).Value = "Defect Location Code Description";
    WSout.Cell(1, 23).Value = "Defect Type Code Group";
    WSout.Cell(1, 24).Value = "Defect Type Code Description";
    WSout.Cell(1, 25).Value = "Cause Type Code Group";
    WSout.Cell(1, 26).Value = "Casue Type Code Description";
    WSout.Cell(1, 27).Value = "Zone";
    WSout.Cell(1, 28).Value = "Sheet";
    WSout.Cell(1, 29).Value = "Nominal";
    WSout.Cell(1, 30).Value = "Toll(-)";
    WSout.Cell(1, 31).Value = "Toll(+)";
    WSout.Cell(1, 32).Value = "Actual";
    WSout.Cell(1, 33).Value = "Extent Of Variation";
    WSout.Cell(1, 34).Value = "Comments Short";
    WSout.Cell(1, 35).Value = "Comments Long";
   // WS.Cell(1, 33).Value = "Extent Of Variation";
   // WS.Cell(1, 33).Value = "Extent Of Variation";

}


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