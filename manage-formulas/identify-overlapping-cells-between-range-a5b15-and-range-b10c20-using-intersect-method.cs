using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsIntersectDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the first range A5:B15
                AsposeRange range1 = cells.CreateRange("A5:B15");

                // Define the second range B10:C20
                AsposeRange range2 = cells.CreateRange("B10:C20");

                // Find the intersection of the two ranges
                AsposeRange intersectRange = range1.Intersect(range2);

                if (intersectRange != null)
                {
                    Console.WriteLine("Intersected range address: " + intersectRange.Address);

                    // Highlight the intersected area
                    Style highlightStyle = workbook.CreateStyle();
                    highlightStyle.ForegroundColor = Color.Yellow;
                    highlightStyle.Pattern = BackgroundType.Solid;
                    intersectRange.SetStyle(highlightStyle);
                }
                else
                {
                    Console.WriteLine("The specified ranges do not intersect.");
                }

                // Save the result workbook
                string outputPath = "IntersectResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}