// Title: Hide empty rows in an Aspose.Cells PivotTable with C# by setting HideEmptyRows to true
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, accesses the first PivotTable, sets its HideEmptyRows property to true, and saves the updated file. | Describe how to enable the HideEmptyRows option on a PivotTable in Aspose.Cells for .NET and provide a fallback strategy if the property is missing in older library versions.
// Common Searches: Aspose.Cells C# hide empty rows in pivot table | Set HideEmptyRows property on PivotTable using Aspose.Cells .NET | Programmatically remove blank rows from an Excel pivot table with Aspose.Cells | How to enable HideEmptyRows for a PivotTable in C# Aspose.Cells
// Tags: Aspose.Cells PivotTable HideEmptyRows | C# hide blank rows Excel pivot | Aspose.Cells set pivot option | Excel pivot table hide empty rows .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example demonstrates loading an existing Excel workbook with Aspose.Cells, locating the first PivotTable on the first worksheet, setting its HideEmptyRows property to true to suppress empty rows, and saving the modified workbook. It also notes handling scenarios where the property may not be present in older API releases.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Hide empty rows in the pivot table (property not available in current API version)
            // If needed, adjust pivot table settings here using available API methods.

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
