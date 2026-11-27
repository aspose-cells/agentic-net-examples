// Title: How to add a column sparkline to cell C5 from range D5:D15 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to create a column sparkline in a single cell | Set data range D5:D15 for a column sparkline with Aspose.Cells | Apply a preset style to a column sparkline in an Aspose.Cells workbook | Save an Excel file that contains a column sparkline using Aspose.Cells for .NET | Example of placing a column sparkline in cell C5 with Aspose.Cells C#
// Tags: add column sparkline Aspose.Cells C# | column sparkline data range D5:D15 | place sparkline in cell C5 worksheet | apply sparkline preset style Aspose.Cells | save workbook as xlsx Aspose.Cells | create sparkline group Aspose.Cells API

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ColumnSparklineDemo
{
    // Demonstrates using Aspose.Cells for .NET (C#) to create a column sparkline, define its source range D5:D15, position it in cell C5, apply a preset style, and save the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the source data range D5:D15 with sample values (1..11)
                Cells cells = sheet.Cells;
                for (int row = 4; row <= 14; row++) // Zero‑based index: row 4 = Excel row 5
                {
                    cells[row, 3].PutValue(row - 3);
                }

                // Define the data range for the sparkline
                string dataRange = "D5:D15";

                // Define the location where the sparkline will be placed (single cell C5)
                CellArea location = CellArea.CreateCellArea("C5", "C5");

                // Add a column sparkline group.
                // Use the SparklineType enum; Column = 1
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, dataRange, false, location);

                // Optional: customize the sparkline appearance (preset style)
                SparklineGroup sparkGroup = sheet.SparklineGroups[groupIndex];
                sparkGroup.PresetStyle = SparklinePresetStyleType.Style1; // any preset style

                // Save the workbook
                string outputPath = "ColumnSparkline.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
