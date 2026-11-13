// Title: Apply a light blue solid fill to a pivot table row field header cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set background color of pivot table row header cell | How to change pivot field header style to light blue using Aspose.Cells | C# code to apply solid fill to first row field header in an Excel pivot table | Using Aspose.Cells to format pivot table header cell color in .NET | Apply light blue fill to pivot table header in a workbook with Aspose.Cells
// Tags: pivot table header cell background Aspose.Cells | set solid fill color C# Aspose.Cells | format pivot row field header .NET | apply light blue style to Excel cell Aspose | modify pivot table cell style using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // Loads an existing workbook, accesses the first pivot table, retrieves the first row field header cell, applies a light blue solid background fill to that cell using Aspose.Cells styling APIs, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
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

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("Pivot table has no row fields.");
                    return;
                }

                // Get the cell that contains the first row field header.
                // The header is located at the start of the pivot table's row range.
                CellArea rowArea = pivotTable.RowRange;
                Cell headerCell = worksheet.Cells[rowArea.StartRow, rowArea.StartColumn];

                // Apply a light blue solid fill to the header cell
                Style headerStyle = headerCell.GetStyle();
                headerStyle.ForegroundColor = Color.LightBlue;
                headerStyle.Pattern = BackgroundType.Solid;
                headerCell.SetStyle(headerStyle);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
