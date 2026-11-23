// Title: Apply LightYellow background to intersected cells of two named ranges and save as XLSX using Aspose.Cells for .NET
// Common Searches: how to highlight overlapping cells of two named ranges using Aspose.Cells C# | Aspose.Cells C# apply solid fill to range intersection | save workbook as xlsx after styling intersected range Aspose.Cells | retrieve intersected range of named ranges Aspose.Cells .NET | set background color for intersected area of named ranges in Aspose.Cells
// Tags: range.Intersect method styling Aspose.Cells | background color for intersected named ranges .NET | solid fill style on overlapping cells Aspose.Cells | save workbook as Xlsx with formatted cells Aspose.Cells | programmatic creation of named ranges Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, defines two overlapping named ranges, obtains their intersected area with the Range.Intersect method, applies a LightYellow solid background style to that intersected range, and saves the result as IntersectedRange.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("1");
                sheet.Cells["B1"].PutValue("2");
                sheet.Cells["C1"].PutValue("3");
                sheet.Cells["D1"].PutValue("4");
                sheet.Cells["A2"].PutValue("5");
                sheet.Cells["B2"].PutValue("6");
                sheet.Cells["C2"].PutValue("7");
                sheet.Cells["D2"].PutValue("8");
                sheet.Cells["A3"].PutValue("9");
                sheet.Cells["B3"].PutValue("10");
                sheet.Cells["C3"].PutValue("11");
                sheet.Cells["D3"].PutValue("12");
                sheet.Cells["A4"].PutValue("13");
                sheet.Cells["B4"].PutValue("14");
                sheet.Cells["C4"].PutValue("15");
                sheet.Cells["D4"].PutValue("16");

                // Define two named ranges
                sheet.Cells.CreateRange("A1:C3").Name = "Range1";
                sheet.Cells.CreateRange("B2:D4").Name = "Range2";

                // Retrieve the named ranges
                Aspose.Cells.Range range1 = workbook.Worksheets.GetRangeByName("Range1");
                Aspose.Cells.Range range2 = workbook.Worksheets.GetRangeByName("Range2");

                // Get the intersected area of the two ranges
                Aspose.Cells.Range intersected = range1.Intersect(range2);

                // Apply a background color to the intersected area
                if (intersected != null && intersected.RowCount > 0 && intersected.ColumnCount > 0)
                {
                    Style style = workbook.CreateStyle();
                    style.ForegroundColor = Color.LightYellow;
                    style.Pattern = BackgroundType.Solid;

                    // Apply the style to the intersected range
                    intersected.ApplyStyle(style, new StyleFlag { CellShading = true });
                }

                // Save the workbook as XLSX
                workbook.Save("IntersectedRange.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
