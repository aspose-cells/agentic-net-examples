using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace IntersectNamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data
                for (int row = 0; row < 6; row++)
                {
                    for (int col = 0; col < 6; col++)
                    {
                        cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define two named ranges:
                //   Range1 -> A2:C4
                //   Range2 -> B3:E5
                int idx1 = workbook.Worksheets.Names.Add("Range1");
                Name name1 = workbook.Worksheets.Names[idx1];
                name1.RefersTo = "=Sheet1!$A$2:$C$4";

                int idx2 = workbook.Worksheets.Names.Add("Range2");
                Name name2 = workbook.Worksheets.Names[idx2];
                name2.RefersTo = "=Sheet1!$B$3:$E$5";

                // Retrieve the actual Range objects from the names
                AsposeRange[] ranges1 = name1.GetRanges(); // should contain one range
                AsposeRange[] ranges2 = name2.GetRanges();

                if (ranges1 == null || ranges1.Length == 0 || ranges2 == null || ranges2.Length == 0)
                {
                    Console.WriteLine("Failed to retrieve named ranges.");
                    return;
                }

                AsposeRange range1 = ranges1[0];
                AsposeRange range2 = ranges2[0];

                // Compute the intersection of the two ranges
                AsposeRange intersectRange = range1.Intersect(range2);

                if (intersectRange != null)
                {
                    // Create a style with yellow fill
                    Style yellowStyle = workbook.CreateStyle();
                    yellowStyle.ForegroundColor = Color.Yellow;
                    yellowStyle.Pattern = BackgroundType.Solid;

                    // Apply the style to the intersected area
                    intersectRange.SetStyle(yellowStyle);

                    Console.WriteLine($"Intersection found: {intersectRange.Address}");
                }
                else
                {
                    Console.WriteLine("The named ranges do not intersect.");
                }

                // Save the workbook
                string outputPath = "IntersectNamedRangesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}