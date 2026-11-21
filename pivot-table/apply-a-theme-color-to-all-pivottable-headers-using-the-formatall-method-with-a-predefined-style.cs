// Title: Apply a solid background color to all PivotTable headers in an Excel workbook using Aspose.Cells C# FormatAll method
// Common Searches: C# Aspose.Cells set background color for all PivotTable headers in a workbook | How to use PivotTable.FormatAll to style header rows with a custom style in Aspose.Cells .NET | Apply a predefined style to every PivotTable header across worksheets with Aspose.Cells | Change PivotTable header color programmatically in Aspose.Cells for .NET | Iterate worksheets and format all PivotTable headers using Aspose.Cells API
// Tags: Aspose.Cells PivotTable FormatAll styling | custom style for PivotTable headers C# | solid background fill for Excel PivotTable cells | iterate worksheets to format PivotTables Aspose.Cells | Excel header color customization with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Pivot;   // PivotTable resides in this namespace

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, creates a solid light‑blue style, loops through each worksheet and its PivotTables, applies the style to all cells (including header rows) via PivotTable.FormatAll, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
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

                // Create a style with a solid background color (fallback when theme colors are unavailable)
                Style headerStyle = workbook.CreateStyle();
                headerStyle.ForegroundColor = Color.LightBlue;   // Use a visible background color
                headerStyle.Pattern = BackgroundType.Solid;      // Solid fill

                // Apply the style to all PivotTable headers in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        try
                        {
                            // Apply the style to all cells of the PivotTable (including headers)
                            pivot.FormatAll(headerStyle);
                        }
                        catch (Exception exPivot)
                        {
                            Console.WriteLine($"Failed to format PivotTable on sheet '{sheet.Name}': {exPivot.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
