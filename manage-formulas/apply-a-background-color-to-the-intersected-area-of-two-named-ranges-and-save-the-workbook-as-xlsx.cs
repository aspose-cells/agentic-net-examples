using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace IntersectNamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (create rule)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue("A1");
                cells["B1"].PutValue("B1");
                cells["C1"].PutValue("C1");
                cells["A2"].PutValue("A2");
                cells["B2"].PutValue("B2");
                cells["C2"].PutValue("C2");
                cells["D2"].PutValue("D2");
                cells["A3"].PutValue("A3");
                cells["B3"].PutValue("B3");
                cells["C3"].PutValue("C3");
                cells["D3"].PutValue("D3");
                cells["B4"].PutValue("B4");
                cells["C4"].PutValue("C4");
                cells["D4"].PutValue("D4");

                // Define two named ranges
                Aspose.Cells.Range firstRange = cells.CreateRange("A1:C3");
                firstRange.Name = "FirstRange";

                Aspose.Cells.Range secondRange = cells.CreateRange("B2:D4");
                secondRange.Name = "SecondRange";

                // Get the intersected area (Range.Intersect method)
                Aspose.Cells.Range intersected = firstRange.Intersect(secondRange);

                if (intersected != null)
                {
                    // Create a style with solid background color
                    Style highlightStyle = workbook.CreateStyle();
                    highlightStyle.ForegroundColor = Color.Yellow;
                    highlightStyle.Pattern = BackgroundType.Solid;

                    // Apply the style to the intersected range
                    intersected.SetStyle(highlightStyle);
                }

                // Save the workbook as XLSX (save rule)
                string outputPath = "IntersectedNamedRanges.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}