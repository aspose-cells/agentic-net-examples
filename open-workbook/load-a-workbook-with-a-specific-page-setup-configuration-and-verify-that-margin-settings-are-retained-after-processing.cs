using System;
using Aspose.Cells;

namespace AsposeCellsMarginDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (replace with provided create rule if available)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set custom margin values (centimeters)
            worksheet.PageSetup.TopMargin = 1.0;      // Top margin
            worksheet.PageSetup.BottomMargin = 1.5;   // Bottom margin
            worksheet.PageSetup.LeftMargin = 0.5;     // Left margin
            worksheet.PageSetup.RightMargin = 0.5;    // Right margin

            // Save the workbook (replace with provided save rule if available)
            workbook.Save("MarginsConfigured.xlsx");

            // Load the workbook back (replace with provided load rule if available)
            Workbook loadedWorkbook = new Workbook("MarginsConfigured.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Verify that margin settings are retained
            Console.WriteLine("Verified Margin Settings after Reload:");
            Console.WriteLine($"Top Margin (cm): {loadedWorksheet.PageSetup.TopMargin}");
            Console.WriteLine($"Bottom Margin (cm): {loadedWorksheet.PageSetup.BottomMargin}");
            Console.WriteLine($"Left Margin (cm): {loadedWorksheet.PageSetup.LeftMargin}");
            Console.WriteLine($"Right Margin (cm): {loadedWorksheet.PageSetup.RightMargin}");

            // Author note: This example demonstrates setting, saving, loading, and verifying page margins using Aspose.Cells.
        }
    }
}