// Title: Copy Worksheet and Verify Paper Size, Margins, and Orientation with Aspose.Cells for .NET
// Description: Demonstrates how to copy a worksheet, duplicate its page‑setup (A3 paper, landscape orientation, custom margins) using Aspose.Cells, and programmatically compare each property to confirm an exact match before saving the workbook.
// Keywords: Aspose.Cells | C# worksheet copy | PageSetup.Copy | compare page setup | paper size A3 | landscape orientation | margin comparison | Excel print settings | validate worksheet copy | Aspose.Cells .NET
// Common Searches: Aspose.Cells copy worksheet page setup | compare worksheet margins C# | verify paper size after Worksheet.Copy | check orientation after copying sheet Aspose.Cells | how to validate print settings in copied Excel sheet
// Developer Intent: Confirm that the page‑setup configuration of a copied worksheet (paper size, orientation, margins) is identical to the source.
// Use Cases: Automated report generation where a template’s print layout must be preserved | Batch duplication of worksheets across workbooks while maintaining exact print settings | Quality‑assurance script that flags mismatched page setup before distribution | Migration of legacy Excel templates to new workbooks with guaranteed layout fidelity
// AI Prompts: Generate C# code using Aspose.Cells that copies a worksheet and then verifies that paper size, orientation, and all margins are identical between source and destination. | Show how to employ PageSetup.Copy together with tolerance checks to confirm exact page layout replication after Worksheet.Copy. | Provide a concise example that logs each page‑setup property comparison result and reports whether all properties match.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupComparison
{
    // Demonstrates how to copy a worksheet, duplicate its page‑setup (A3 paper, landscape orientation, custom margins) using Aspose.Cells, and programmatically compare each property to confirm an exact match before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create source workbook and configure its page setup
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";

            // Set page setup properties
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Landscape;
            sourceSheet.PageSetup.TopMargin = 1.5;      // centimeters
            sourceSheet.PageSetup.BottomMargin = 2.0;   // centimeters
            sourceSheet.PageSetup.LeftMargin = 1.0;     // centimeters
            sourceSheet.PageSetup.RightMargin = 1.0;    // centimeters

            // Add some data to the source sheet (optional)
            sourceSheet.Cells["A1"].PutValue("Sample Data");

            // Create destination workbook
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "DestinationSheet";

            // Copy contents and formats from source sheet to destination sheet
            destSheet.Copy(sourceSheet);

            // Copy page setup settings using the provided PageSetup.Copy method
            destSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Compare paper size, orientation, and margins between source and destination
            bool paperSizeMatch = destSheet.PageSetup.PaperSize == sourceSheet.PageSetup.PaperSize;
            bool orientationMatch = destSheet.PageSetup.Orientation == sourceSheet.PageSetup.Orientation;
            bool topMarginMatch = Math.Abs(destSheet.PageSetup.TopMargin - sourceSheet.PageSetup.TopMargin) < 0.0001;
            bool bottomMarginMatch = Math.Abs(destSheet.PageSetup.BottomMargin - sourceSheet.PageSetup.BottomMargin) < 0.0001;
            bool leftMarginMatch = Math.Abs(destSheet.PageSetup.LeftMargin - sourceSheet.PageSetup.LeftMargin) < 0.0001;
            bool rightMarginMatch = Math.Abs(destSheet.PageSetup.RightMargin - sourceSheet.PageSetup.RightMargin) < 0.0001;

            bool allMatch = paperSizeMatch && orientationMatch && topMarginMatch &&
                            bottomMarginMatch && leftMarginMatch && rightMarginMatch;

            Console.WriteLine("Paper size match: " + paperSizeMatch);
            Console.WriteLine("Orientation match: " + orientationMatch);
            Console.WriteLine("Top margin match: " + topMarginMatch);
            Console.WriteLine("Bottom margin match: " + bottomMarginMatch);
            Console.WriteLine("Left margin match: " + leftMarginMatch);
            Console.WriteLine("Right margin match: " + rightMarginMatch);
            Console.WriteLine("All page setup properties match: " + allMatch);

            // Save the destination workbook
            destWorkbook.Save("DestinationWorkbook.xlsx");
        }
    }
}
