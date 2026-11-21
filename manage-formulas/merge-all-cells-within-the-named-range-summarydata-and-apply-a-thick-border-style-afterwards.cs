// Title: Merge a named range in an Excel workbook and apply a thick border style using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge cells in a named range and set thick borders | How to apply only border formatting to a merged range with Aspose.Cells .NET | Aspose.Cells .NET create thick border style and apply to a named range | Merge range by name and add border style using Aspose.Cells in C# | Apply StyleFlag borders to merged cells in an Aspose.Cells workbook
// Tags: merge named range Aspose.Cells .NET | apply thick border style Aspose.Cells | StyleFlag borders C# | range.ApplyStyle merged cells | Excel workbook cell formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, retrieves the named range "SummaryData", merges all its cells, creates a thick border style, applies only the border formatting to the merged range using a StyleFlag, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Retrieve the named range "SummaryData"
                var range = workbook.Worksheets.GetRangeByName("SummaryData");
                if (range == null)
                {
                    Console.WriteLine("Named range 'SummaryData' not found.");
                    return;
                }

                // Merge all cells within the named range
                range.Merge();

                // Create a style with thick borders on all sides
                var thickBorderStyle = workbook.CreateStyle();
                thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

                // Apply only the border style to the merged range
                var styleFlag = new StyleFlag { Borders = true };
                range.ApplyStyle(thickBorderStyle, styleFlag);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
