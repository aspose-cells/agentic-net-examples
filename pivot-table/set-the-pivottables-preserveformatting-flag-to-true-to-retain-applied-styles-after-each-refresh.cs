// Title: How to Preserve PivotTable Formatting in Aspose.Cells for .NET (C#)
// Description: Demonstrates loading an existing workbook, accessing its first PivotTable, enabling the PreserveFormatting flag, applying a custom style, refreshing the data, and saving the file so that all formatting remains intact after each refresh.
// Keywords: Aspose.Cells PivotTable PreserveFormatting | C# keep pivot formatting after refresh | Aspose.Cells FormatAll example | Excel pivot table style retention .NET | RefreshData PreserveFormatting Aspose | Aspose.Cells pivot table formatting | retain pivot table style C# | Aspose.Cells workbook automation | Excel automation PreserveFormatting | Aspose.Cells PivotTable example
// Common Searches: Aspose.Cells PreserveFormatting property | keep pivot table formatting after RefreshData | C# Aspose.Cells apply style to PivotTable | how to retain Excel pivot formatting with code | Aspose.Cells PivotTable FormatAll usage | preserve custom pivot styles in .NET
// Developer Intent: Enable the PreserveFormatting flag on a PivotTable so that any custom styles stay applied when the table is refreshed programmatically.
// Use Cases: Maintain font, color, and border settings on a pivot report that is refreshed daily. | Ensure conditional formatting rules survive data updates in automated Excel dashboards. | Apply a corporate style once and have it persist across multiple programmatic refresh cycles.
// AI Prompts: Provide C# code using Aspose.Cells that loads a workbook, sets PivotTable.PreserveFormatting = true, applies a custom style, refreshes the data, and saves the file. | Show an example that loops through all PivotTables in a workbook, enables PreserveFormatting for each, refreshes them, and writes the result. | Explain the interaction between PreserveFormatting, FormatAll, and RefreshData in Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates loading an existing workbook, accessing its first PivotTable, enabling the PreserveFormatting flag, applying a custom style, refreshing the data, and saving the file so that all formatting remains intact after each refresh.
    public class PivotTablePreserveFormattingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"The required input file '{inputPath}' was not found.");
            }

            // Load the existing workbook that contains a pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                throw new InvalidOperationException("No pivot tables found in the first worksheet.");
            }

            // Access the first pivot table in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable preserving formatting when the pivot table is refreshed
            pivotTable.PreserveFormatting = true;

            // OPTIONAL: Apply a style to the pivot table data area to demonstrate preservation
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            pivotTable.FormatAll(style);

            // Refresh the pivot table data (formatting will be kept because PreserveFormatting is true)
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
