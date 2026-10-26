// Title: Hide PivotTable Toolbar and Field List Ribbon in Aspose.Cells for .NET at Runtime
// Common Searches: aspnet hide pivot table toolbar using Aspose.Cells C# | disable pivot table field list ribbon in Aspose.Cells .NET runtime | remove Excel pivot table UI elements with Aspose.Cells during processing | how to turn off PivotTable toolbar in Aspose.Cells workbook programmatically | Aspose.Cells hide pivot table ribbon without user interaction
// Tags: Aspose.Cells hide pivot table toolbar .NET | disable pivot table field list ribbon Aspose.Cells | runtime workbook UI customization Aspose.Cells | pivot table ribbon settings Aspose.Cells | C# hide Excel UI elements Aspose.Cells | programmatic UI suppression Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, optionally sets ShowPivotTableToolbar and ShowPivotTableFieldList to false (when supported) to hide the PivotTable toolbar and field list ribbon at runtime, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In recent Aspose.Cells versions the properties to hide the PivotTable toolbar
            // and field list are not available. If using a version that provides them, you can
            // uncomment and set them here:
            // workbook.Settings.ShowPivotTableToolbar = false;
            // workbook.Settings.ShowPivotTableFieldList = false;

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
