// Title: Compare paper size, width, and height of two worksheets using Aspose.Cells for .NET
// Description: Loads an Excel workbook, selects two worksheets by name, reads each sheet's PageSetup, compares PaperSize, PaperWidth, and PaperHeight, logs any differences to the console, and saves the workbook unchanged.
// Keywords: Aspose.Cells | C# | .NET | worksheet page setup comparison | paper size difference | PaperWidth | PaperHeight | compare Excel sheets | print layout validation | PageSetup API
// Common Searches: Aspose.Cells compare worksheet page setup | C# detect different paper size between Excel sheets | log paper dimensions differences Aspose.Cells | check print layout consistency across worksheets .NET | compare sheet print settings with Aspose.Cells
// Developer Intent: Identify mismatches in paper size, width, or height between two specified worksheets and output the discrepancies.
// Use Cases: Validate that all sheets in a reporting workbook share identical print settings before bulk printing. | Audit template worksheets to ensure consistent page layout for automated document generation. | Generate a console report of page‑setup differences for quality‑control checks in a CI pipeline.
// AI Prompts: Write a reusable method that returns a list of PageSetup property differences between two worksheets using Aspose.Cells. | Extend the comparison to include orientation, margins, and scaling, and output the results to a structured log file. | Refactor the logic into a utility class with separate functions for paper size, width, height, and other page‑setup attributes.

using System;
using Aspose.Cells;

namespace PaperDimensionComparer
{
    // Loads an Excel workbook, selects two worksheets by name, reads each sheet's PageSetup, compares PaperSize, PaperWidth, and PaperHeight, logs any differences to the console, and saves the workbook unchanged.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file containing the worksheets to compare
            string workbookPath = "InputWorkbook.xlsx";

            // Names of the two worksheets to compare
            string firstSheetName = "Sheet1";
            string secondSheetName = "Sheet2";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the worksheets
            Worksheet sheet1 = workbook.Worksheets[firstSheetName];
            Worksheet sheet2 = workbook.Worksheets[secondSheetName];

            if (sheet1 == null || sheet2 == null)
            {
                Console.WriteLine("One or both specified worksheets were not found.");
                return;
            }

            // Access PageSetup for each sheet
            PageSetup ps1 = sheet1.PageSetup;
            PageSetup ps2 = sheet2.PageSetup;

            // Compare PaperSize enum
            if (ps1.PaperSize != ps2.PaperSize)
            {
                Console.WriteLine($"PaperSize differs: {firstSheetName} = {ps1.PaperSize}, {secondSheetName} = {ps2.PaperSize}");
            }

            // Compare PaperWidth (in inches, orientation considered)
            double width1 = ps1.PaperWidth;
            double width2 = ps2.PaperWidth;
            if (Math.Abs(width1 - width2) > 0.001) // tolerance for floating point
            {
                Console.WriteLine($"PaperWidth differs: {firstSheetName} = {width1:F3} in, {secondSheetName} = {width2:F3} in");
            }

            // Compare PaperHeight (in inches, orientation considered)
            double height1 = ps1.PaperHeight;
            double height2 = ps2.PaperHeight;
            if (Math.Abs(height1 - height2) > 0.001)
            {
                Console.WriteLine($"PaperHeight differs: {firstSheetName} = {height1:F3} in, {secondSheetName} = {height2:F3} in");
            }

            // If no differences were found
            if (ps1.PaperSize == ps2.PaperSize && Math.Abs(width1 - width2) <= 0.001 && Math.Abs(height1 - height2) <= 0.001)
            {
                Console.WriteLine("No differences in paper dimensions between the two worksheets.");
            }

            // Optionally, save the workbook after any modifications (save rule)
            // In this example we do not modify the workbook, but the save call demonstrates the rule usage.
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
