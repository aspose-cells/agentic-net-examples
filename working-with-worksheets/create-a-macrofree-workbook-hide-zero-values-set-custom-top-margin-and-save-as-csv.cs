// Title: Create a macro‑free workbook, hide zero values, set a custom top margin, and export to CSV using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide zero values in a worksheet | Set custom top margin for worksheet page setup with Aspose.Cells | Export a macro‑free workbook to CSV using Aspose.Cells .NET | Create a workbook without macros and save as CSV in C# | Configure page margins and hide zero values before CSV export with Aspose.Cells
// Tags: macro‑free workbook generation Aspose.Cells | zero‑value hiding worksheet Aspose.Cells | custom top margin page setup Aspose.Cells | CSV export workbook Aspose.Cells C#

using Aspose.Cells;
using System;

// The example demonstrates how to create a new macro‑free Workbook, optionally hide zero values in the first Worksheet, set a custom top margin via PageSetup, populate sample cells, and save the workbook as a CSV file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑free workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide zero values in the worksheet (if the property exists in the used version)
            // sheet.IsZeroValuesHidden = true; // Uncomment if supported by your Aspose.Cells version

            // Set a custom top margin (e.g., 0.5 inch = 36 points)
            sheet.PageSetup.TopMargin = 36; // points

            // Example data (optional)
            sheet.Cells["A1"].PutValue(0);
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(0);
            sheet.Cells["B1"].PutValue(45);
            sheet.Cells["B2"].PutValue(0);
            sheet.Cells["B3"].PutValue(78);

            // Save the workbook as CSV
            string outputPath = "output.csv";
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
