// Title: Aspose.Cells C# – Compare and Log Worksheet Paper Size, Width & Height
// Description: Load an Excel workbook with Aspose.Cells, retrieve two worksheets, compare their PageSetup properties (PaperSize, PaperWidth, PaperHeight), display differences in the console, and save a comparison log workbook (PaperDimensionComparisonLog.xlsx).
// Keywords: Aspose.Cells | C# compare worksheet page setup | paper size difference | PaperWidth | PaperHeight | PageSetup comparison | Excel log workbook | worksheet dimensions | print layout validation | Aspose.Cells API
// Common Searches: Aspose.Cells compare worksheet paper size | C# compare Excel sheet PageSetup | log page setup differences Aspose | save worksheet comparison report Aspose.Cells | how to check PaperWidth in Aspose.Cells
// Developer Intent: Detect mismatched page‑setup settings (paper size, width, height) between two worksheets and generate a reusable log file.
// Use Cases: Ensure printed reports from multiple sheets share identical page settings before distribution. | Create a discrepancy report when consolidating workbooks from different sources to maintain uniform print layout. | Integrate a quality‑check step in CI pipelines that flags unintended changes to worksheet page‑setup properties.
// AI Prompts: Write a reusable C# method that takes two worksheet names and returns a dictionary of PageSetup properties that differ, using Aspose.Cells. | Show how to extend the comparison to include orientation, margins, and header/footer settings, and export the results to a CSV file. | Generate C# code that iterates over all worksheets in a workbook, compares their PageSetup configurations, and summarizes the differences in a single Excel report.

using System;
using Aspose.Cells;

namespace PaperDimensionComparer
{
    // Load an Excel workbook with Aspose.Cells, retrieve two worksheets, compare their PageSetup properties (PaperSize, PaperWidth, PaperHeight), display differences in the console, and save a comparison log workbook (PaperDimensionComparisonLog.xlsx).
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "input.xlsx";

            // Names of the two worksheets to compare
            string sheetName1 = "Sheet1";
            string sheetName2 = "Sheet2";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(filePath);

            // Retrieve the worksheets
            Worksheet ws1 = workbook.Worksheets[sheetName1];
            Worksheet ws2 = workbook.Worksheets[sheetName2];

            if (ws1 == null || ws2 == null)
            {
                Console.WriteLine("One or both specified worksheets were not found.");
                return;
            }

            // Access PageSetup for each worksheet
            PageSetup ps1 = ws1.PageSetup;
            PageSetup ps2 = ws2.PageSetup;

            // Compare PaperSize enum
            if (ps1.PaperSize != ps2.PaperSize)
            {
                Console.WriteLine($"PaperSize differs: {sheetName1} = {ps1.PaperSize}, {sheetName2} = {ps2.PaperSize}");
            }
            else
            {
                Console.WriteLine($"PaperSize is the same: {ps1.PaperSize}");
            }

            // Compare PaperWidth (in inches)
            double width1 = ps1.PaperWidth;
            double width2 = ps2.PaperWidth;
            if (Math.Abs(width1 - width2) > 0.0001)
            {
                Console.WriteLine($"PaperWidth differs: {sheetName1} = {width1} in, {sheetName2} = {width2} in");
            }
            else
            {
                Console.WriteLine($"PaperWidth is the same: {width1} in");
            }

            // Compare PaperHeight (in inches)
            double height1 = ps1.PaperHeight;
            double height2 = ps2.PaperHeight;
            if (Math.Abs(height1 - height2) > 0.0001)
            {
                Console.WriteLine($"PaperHeight differs: {sheetName1} = {height1} in, {sheetName2} = {height2} in");
            }
            else
            {
                Console.WriteLine($"PaperHeight is the same: {height1} in");
            }

            // Optionally, save a log workbook (save rule)
            Workbook logWorkbook = new Workbook();
            Worksheet logSheet = logWorkbook.Worksheets[0];
            logSheet.Name = "ComparisonLog";

            int row = 0;
            logSheet.Cells[row, 0].PutValue("Property");
            logSheet.Cells[row, 1].PutValue(sheetName1);
            logSheet.Cells[row, 2].PutValue(sheetName2);
            row++;

            // Log PaperSize
            logSheet.Cells[row, 0].PutValue("PaperSize");
            logSheet.Cells[row, 1].PutValue(ps1.PaperSize.ToString());
            logSheet.Cells[row, 2].PutValue(ps2.PaperSize.ToString());
            row++;

            // Log PaperWidth
            logSheet.Cells[row, 0].PutValue("PaperWidth (in)");
            logSheet.Cells[row, 1].PutValue(width1);
            logSheet.Cells[row, 2].PutValue(width2);
            row++;

            // Log PaperHeight
            logSheet.Cells[row, 0].PutValue("PaperHeight (in)");
            logSheet.Cells[row, 1].PutValue(height1);
            logSheet.Cells[row, 2].PutValue(height2);
            row++;

            // Save the log workbook
            string logPath = "PaperDimensionComparisonLog.xlsx";
            logWorkbook.Save(logPath);
            Console.WriteLine($"Comparison log saved to {logPath}");
        }
    }
}
