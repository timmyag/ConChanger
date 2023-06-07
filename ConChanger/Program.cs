using ClosedXML.Excel;

internal class Program
{
    static class Globals
    {
        public static string FilePath = "\"out.XLSX\""  ;
    }
        private static void Main(string[] args)
    {
         
    Console.WriteLine("Welcome to the ConConverter");

        // The Con workbook beeing read
        string InPath = "C:\\Users\\tguiv\\OneDrive\\Desktop\\New folder\\DAS 18639.XLSX"; //default file for debugging
        if (args.Count() != 0) //takes a file that is droped onto the .exe, should be the default during normal use
        {
            Console.WriteLine("Opening " + args[0].ToString());
            InPath = args[0].ToString();
        }
        Body(InPath);
    }
    private static void Body(string InPath)
    {
        Console.WriteLine("Opening the Concession XLSX ");
        XLWorkbook wbIn = new XLWorkbook(new FileStream(InPath, FileMode.Open, FileAccess.Read, FileShare.Read));

        int ConLineCount = CountLinesInCon(wbIn);

        Console.WriteLine("Opened the Concession and found " + ConLineCount.ToString() + " lines");

        var wbOut = GetWorkBook(); //get refference to the out workbook
        var wbOutLineCount = CountLinesOutBook(wbOut);

        Headers(wbOut);
        ConSheet(wbIn, wbOut, ConLineCount);
        ContinuationSheet(wbIn, wbOut, ConLineCount);

        Console.WriteLine("Done with the Data, saving the file");
        wbOut.SaveAs(Globals.FilePath);

        Console.WriteLine("Done");

        // leave the terminal window open for 3s
        Thread.Sleep(3000);
    }

    private static int CountLinesOutBook(XLWorkbook wbOut)
    {
        throw new NotImplementedException();
    }

    private static XLWorkbook GetWorkBook()
    {
        if (File.Exists(Globals.FilePath)) 
            {return new XLWorkbook(Globals.FilePath);}
        else
        {
            var wbOut = new XLWorkbook();
            wbOut.AddWorksheet();
            return wbOut;
        }
    }

        // Moving the data off the continuation sheet
        static void ContinuationSheet(XLWorkbook wbin, XLWorkbook wbout, int ConLineCount)
        {
            IXLWorksheet WSout = wbout.Worksheets.First();
            IXLWorksheet ContSheet = wbin.Worksheet(3);
            string ContSheetName = "Continuation Sheet";
            if (ContSheet.Name != ContSheetName) { throw new ArgumentException("The continuation sheet wasnt found"); }

            ContSheet.Range(4, 2, ConLineCount + 3, 2).CopyTo(WSout.Cell(2, 20)); // Item No.
            ContSheet.Range(4, 3, ConLineCount + 3, 3).CopyTo(WSout.Cell(2, 21));// Defect Location Code Group
            ContSheet.Range(4, 6, ConLineCount + 3, 6).CopyTo(WSout.Cell(2, 22)); // Defect Location Code Description
            ContSheet.Range(4, 7, ConLineCount + 3, 7).CopyTo(WSout.Cell(2, 23)); // Defect Type Code Group
            ContSheet.Range(4, 9, ConLineCount + 3, 9).CopyTo(WSout.Cell(2, 24)); //"Defect Type Code Description"
            ContSheet.Range(4, 10, ConLineCount + 3, 10).CopyTo(WSout.Cell(2, 25)); //"Cause Type Code Group"
            ContSheet.Range(4, 12, ConLineCount + 3, 12).CopyTo(WSout.Cell(2, 26)); //"Casue Type Code Description"
            ContSheet.Range(4, 15, ConLineCount + 3, 15).CopyTo(WSout.Cell(2, 27)); //"Zone"
            ContSheet.Range(4, 16, ConLineCount + 3, 16).CopyTo(WSout.Cell(2, 28)); //"Sheet"
            ContSheet.Range(4, 17, ConLineCount + 3, 17).CopyTo(WSout.Cell(2, 29)); //"Nominal"
            ContSheet.Range(4, 18, ConLineCount + 3, 18).CopyTo(WSout.Cell(2, 30)); //"Toll(-)"
            ContSheet.Range(4, 19, ConLineCount + 3, 19).CopyTo(WSout.Cell(2, 31)); //Toll(+)
            ContSheet.Range(4, 20, ConLineCount + 3, 20).CopyTo(WSout.Cell(2, 32)); //Actual
            ContSheet.Range(4, 21, ConLineCount + 3, 21).CopyTo(WSout.Cell(2, 33)); //"Extent Of Variation"
            ContSheet.Range(4, 22, ConLineCount + 3, 22).CopyTo(WSout.Cell(2, 34)); //"Comments Short"
            ContSheet.Range(4, 24, ConLineCount + 3, 24).CopyTo(WSout.Cell(2, 35)); //"Comments Long"

            //clear out the styling mess that came acorss with the abouve
            WSout.ConditionalFormats.RemoveAll();
            foreach (IXLDataValidation val in WSout.DataValidations) { val.Clear(); };
            WSout.Range(2, 20, ConLineCount + 5, 35).Style = WSout.Cell(1, 1).Style;
        }

        // moving the data off the ConDP sheet
        static void ConSheet(XLWorkbook wb, XLWorkbook wbout, int ConLineCount)
        {
            IXLWorksheet WSout = wbout.Worksheets.First();
            IXLWorksheet ConSheet = wb.Worksheet(2);
            string ConSheetName = "Concession or Deviation Permit";
            if (ConSheet.Name != ConSheetName) { throw new ArgumentException("The Con or DP sheet wasnt found"); }

            WSout.Range(2, 1, ConLineCount + 1, 1).Value = ConSheet.Cell(4, 4).Value; //1.1
            WSout.Range(2, 2, ConLineCount + 1, 2).Value = ConSheet.Cell(6, 3).Value; //1.2
            WSout.Range(2, 3, ConLineCount + 1, 3).Value = ConSheet.Cell(6, 8).Value; //1.2
            WSout.Range(2, 4, ConLineCount + 1, 4).Value = ConSheet.Cell(4, 19).Value; //1.3
            WSout.Range(2, 5, ConLineCount + 1, 5).Value = ConSheet.Cell(6, 20).Value; //1.4
            WSout.Range(2, 6, ConLineCount + 1, 6).Value = ConSheet.Cell(9, 1).Value; //2.1
            WSout.Range(2, 7, ConLineCount + 1, 7).Value = ConSheet.Cell(9, 6).Value; //2.2
            WSout.Range(2, 8, ConLineCount + 1, 8).Value = ConSheet.Cell(11, 6).Value; //2.3
            WSout.Range(2, 9, ConLineCount + 1, 9).Value = ConSheet.Cell(9, 16).Value; //2.4
            WSout.Range(2, 10, ConLineCount + 1, 10).Value = ConSheet.Cell(9, 21).Value; //2.5
            WSout.Range(2, 11, ConLineCount + 1, 11).Value = ConSheet.Cell(11, 16).Value; //2.6
            WSout.Range(2, 12, ConLineCount + 1, 12).Value = ConSheet.Cell(11, 19).Value; //2.7
            WSout.Range(2, 13, ConLineCount + 1, 13).Value = ConSheet.Cell(9, 23).Value; //2.8
            WSout.Range(2, 14, ConLineCount + 1, 14).Value = ConSheet.Cell(11, 23).Value; //2.9
            WSout.Range(2, 15, ConLineCount + 1, 15).Value = ConSheet.Cell(24, 1).Value; //Previous Subs of a similar
            WSout.Range(2, 16, ConLineCount + 1, 16).Value = ConSheet.Cell(24, 11).Value; //"Previous Subs to actual"
            WSout.Range(2, 17, ConLineCount + 1, 17).Value = ConSheet.Cell(26, 1).Value;
            WSout.Range(2, 18, ConLineCount + 1, 18).Value = ConSheet.Cell(36, 2).Value;
        }

        // Filling in the headers
        static void Headers(XLWorkbook wbout)
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
            WSout.Cell(1, 18).Value = "Originator";

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

        // count the number of lines in the concession
        static int CountLinesInCon(XLWorkbook wb)
        {
            string ContSheetName = "Continuation Sheet";
            IXLWorksheet ContSheet = wb.Worksheet(3);
            if (ContSheet.Name != ContSheetName) { throw new ArgumentException("The continuation sheet wasnt found"); }

            for (int i = 4; i < 100; i++)
            {
                if (ContSheet.Cell(i, 2).Value.IsBlank)
                { return i - 4; }
            }

            return 0;
        }
    
}