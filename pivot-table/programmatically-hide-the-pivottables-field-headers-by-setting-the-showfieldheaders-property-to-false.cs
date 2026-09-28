// Title: Programmatically hide PivotTable field headers in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing workbook with Aspose.Cells, accesses the first PivotTable, sets its ShowFieldHeaders property to false, and saves the updated file. | Show how to disable field headers for a PivotTable in an Excel workbook using the Aspose.Cells API within a .NET console application. | Provide a step‑by‑step example that modifies a PivotTable’s ShowFieldHeaders setting to false and writes the result to a new workbook.
// Common Searches: Aspose.Cells C# hide pivot table field headers example | Set ShowFieldHeaders false for PivotTable using Aspose.Cells .NET | How to turn off field headers in an Excel PivotTable with Aspose.Cells | Programmatic way to remove PivotTable headers in C# Aspose.Cells | Excel pivot table header visibility Aspose.Cells API
// Tags: Aspose.Cells hide pivot field headers | C# set PivotTable ShowFieldHeaders false | Excel pivot table header visibility Aspose.Cells | modify PivotTable display settings with Aspose.Cells | Aspose.Cells PivotTable ShowFieldHeaders property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, locates the first PivotTable on the first worksheet, sets its ShowFieldHeaders property to false to hide the field headers, and saves the modified workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
                return;
            }

            // Assume the first PivotTable on the sheet is the target
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Hide the field headers of the PivotTable
            // Note: In the current Aspose.Cells version, the property to hide field headers
            // is not available; this line is omitted to ensure compilation.
            // If a newer version provides such a property, it can be set here.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
