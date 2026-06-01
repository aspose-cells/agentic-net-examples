using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupComparison
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create source workbook and configure PageSetup --------------------
            Workbook sourceWorkbook = new Workbook();                     // create a new workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];        // get the first worksheet

            // Set some sample data (optional, just to have content)
            sourceSheet.Cells["A1"].PutValue("Source Sheet");

            // Configure page setup for the source worksheet
            PageSetup srcSetup = sourceSheet.PageSetup;
            srcSetup.PaperSize = PaperSizeType.PaperA4;                 // set paper size
            srcSetup.Orientation = PageOrientationType.Landscape;       // set orientation
            srcSetup.TopMargin = 1.5;                                   // centimeters
            srcSetup.BottomMargin = 1.0;
            srcSetup.LeftMargin = 0.8;
            srcSetup.RightMargin = 0.8;

            // -------------------- Create destination workbook --------------------
            Workbook destWorkbook = new Workbook();                       // create a new (empty) workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];            // get its first worksheet

            // -------------------- Copy source worksheet to destination worksheet --------------------
            // Use the Worksheet.Copy method as defined in the documentation
            destSheet.Copy(sourceSheet);                                 // copies contents and formats

            // -------------------- Compare PageSetup properties --------------------
            PageSetup destSetup = destSheet.PageSetup;

            bool paperSizeMatch = srcSetup.PaperSize == destSetup.PaperSize;
            bool orientationMatch = srcSetup.Orientation == destSetup.Orientation;
            bool topMarginMatch = Math.Abs(srcSetup.TopMargin - destSetup.TopMargin) < 0.0001;
            bool bottomMarginMatch = Math.Abs(srcSetup.BottomMargin - destSetup.BottomMargin) < 0.0001;
            bool leftMarginMatch = Math.Abs(srcSetup.LeftMargin - destSetup.LeftMargin) < 0.0001;
            bool rightMarginMatch = Math.Abs(srcSetup.RightMargin - destSetup.RightMargin) < 0.0001;

            // Output comparison results
            Console.WriteLine("Paper Size match: " + paperSizeMatch);
            Console.WriteLine("Orientation match: " + orientationMatch);
            Console.WriteLine("Top Margin match: " + topMarginMatch);
            Console.WriteLine("Bottom Margin match: " + bottomMarginMatch);
            Console.WriteLine("Left Margin match: " + leftMarginMatch);
            Console.WriteLine("Right Margin match: " + rightMarginMatch);

            // -------------------- Save the destination workbook --------------------
            destWorkbook.Save("DestinationWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}