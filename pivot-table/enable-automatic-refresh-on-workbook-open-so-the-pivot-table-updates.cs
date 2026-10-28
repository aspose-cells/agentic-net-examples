// Title: How to programmatically refresh Excel pivot tables on workbook open using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh all pivot tables automatically when an Excel file is opened with Aspose.Cells | Aspose.Cells .NET refresh pivot table data on workbook load example | How to use Aspose.Cells to update pivot tables during workbook opening in C# | Programmatic pivot table refresh on Excel workbook open using Aspose.Cells for .NET | RefreshData method for Excel pivot tables with Aspose.Cells when opening a file
// Tags: Aspose.Cells pivot RefreshData method usage | C# refresh Excel pivot tables on workbook load | programmatic pivot table update with Aspose.Cells | iterate worksheets and pivot tables Aspose.Cells | automatic pivot refresh on Excel file open .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, iterates through each worksheet and its PivotTable collection, calls PivotTable.RefreshData() to ensure pivot data is current (since RefreshDataOnOpen isn’t available), and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh each pivot table's data to ensure it is up‑to‑date
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        pivot.RefreshData(); // Refresh data now (property RefreshDataOnOpen not available)
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to refresh pivot table '{pivot.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
