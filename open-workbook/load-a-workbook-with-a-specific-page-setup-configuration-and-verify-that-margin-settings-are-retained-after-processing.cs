// Title: C# – Load an Excel workbook with Aspose.Cells and verify that page setup margin settings are retained
// Common Searches: how to read and compare Excel page setup margins with Aspose.Cells in C# | verify that top, bottom, left, right margins are unchanged after loading a workbook using Aspose.Cells .NET | C# example for checking worksheet margin values from an existing .xlsx file with Aspose.Cells | ensure margin settings are retained when saving an Excel file with Aspose.Cells | Aspose.Cells page setup margin validation code snippet for .NET
// Tags: Aspose.Cells page setup margin extraction | C# verify Excel worksheet margins | read margin properties from .xlsx using Aspose.Cells | validate retained page margins Aspose.Cells | Excel margin verification .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates or loads an Excel workbook, accesses the first worksheet's PageSetup, reads the top, bottom, left, and right margin values, compares them against expected inches, throws an exception if any margin differs, and finally saves the workbook. It demonstrates how to ensure margin settings are preserved when opening and saving files with Aspose.Cells for .NET, with comprehensive exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook that should have page setup margins configured
            string sourcePath = "InputWorkbook.xlsx";

            // Ensure the source file exists; if not, create a workbook with known margin settings
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                var tempSetup = tempSheet.PageSetup;
                tempSetup.TopMargin = 0.75;
                tempSetup.BottomMargin = 0.75;
                tempSetup.LeftMargin = 0.7;
                tempSetup.RightMargin = 0.7;
                tempWb.Save(sourcePath);
                Console.WriteLine($"Created placeholder workbook '{sourcePath}' with predefined margins.");
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the page setup object
            PageSetup pageSetup = sheet.PageSetup;

            // Expected margin values (in inches) – match the values set above
            double expectedTopMargin = 0.75;
            double expectedBottomMargin = 0.75;
            double expectedLeftMargin = 0.7;
            double expectedRightMargin = 0.7;

            // Verify that the margins are retained after loading
            bool marginsMatch =
                Math.Abs(pageSetup.TopMargin - expectedTopMargin) < 0.001 &&
                Math.Abs(pageSetup.BottomMargin - expectedBottomMargin) < 0.001 &&
                Math.Abs(pageSetup.LeftMargin - expectedLeftMargin) < 0.001 &&
                Math.Abs(pageSetup.RightMargin - expectedRightMargin) < 0.001;

            if (!marginsMatch)
            {
                Console.WriteLine("Margin verification failed.");
                Console.WriteLine($"Top: {pageSetup.TopMargin}, Expected: {expectedTopMargin}");
                Console.WriteLine($"Bottom: {pageSetup.BottomMargin}, Expected: {expectedBottomMargin}");
                Console.WriteLine($"Left: {pageSetup.LeftMargin}, Expected: {expectedLeftMargin}");
                Console.WriteLine($"Right: {pageSetup.RightMargin}, Expected: {expectedRightMargin}");
                throw new InvalidOperationException("Page setup margins are not as expected.");
            }
            else
            {
                Console.WriteLine("All margin settings are correctly retained.");
            }

            // Optionally, save the workbook after verification (e.g., to a new file)
            string outputPath = "VerifiedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (InvalidOperationException invEx)
        {
            Console.WriteLine($"Invalid operation: {invEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
