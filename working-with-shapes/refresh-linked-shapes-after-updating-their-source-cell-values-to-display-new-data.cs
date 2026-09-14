// Title: How to refresh linked shapes (charts and pictures) after modifying source cells using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, change values in specific cells, and invoke Aspose.Cells to trigger formula calculation so that all linked charts and pictures update automatically. | Write C# code that updates cells A1, A2, and A3 in a worksheet and programmatically refreshes linked shapes without reopening the file. | Demonstrate using Aspose.Cells to modify cell data, recalculate the workbook, and save the file with refreshed linked objects.
// Common Searches: Aspose.Cells C# refresh chart after updating cell values programmatically | How to recalculate formulas and update linked pictures in an Excel file using Aspose.Cells .NET | Programmatically trigger linked shape refresh in Aspose.Cells after changing source data | C# example to update cells and refresh all linked objects in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells refresh chart objects | Aspose.Cells refresh picture objects | Aspose.Cells trigger formula calculation | C# programmatically update Excel cells | Aspose.Cells modify cell data source

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), updates cells A1, A2, and A3 with new values, triggers a workbook-wide formula calculation to refresh any linked charts or pictures, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optionally add a default worksheet
                workbook.Worksheets.Add("Sheet1");
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Update source cells that linked shapes depend on
            sheet.Cells["A1"].PutValue(150);
            sheet.Cells["A2"].PutValue(250);
            sheet.Cells["A3"].PutValue(350);

            // Recalculate formulas (charts will reflect updated data automatically)
            workbook.CalculateFormula();

            // Save the workbook with refreshed data
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
