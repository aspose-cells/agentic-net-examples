// Title: Set the default row height to 20 points in an Aspose.Cells worksheet using C#
// Common Searches: Aspose.Cells C# set row height to 20 points for first row | How to change default row height in an Aspose.Cells workbook with .NET | C# example for adjusting worksheet row height for better readability using Aspose.Cells | Set specific row height in Aspose.Cells when DefaultRowHeight property is unavailable | Increase Excel row height to 20 points with Aspose.Cells API in C#
// Tags: Aspose.Cells set row height C# | worksheet row height adjustment Aspose.Cells | C# Excel row height 20 points | Aspose.Cells Cells.SetRowHeight usage | increase row height readability Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // This C# example creates a new Workbook, accesses the first Worksheet, and uses Cells.SetRowHeight to set the first row's height to 20 points for improved readability, then saves the workbook as output.xlsx while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the height of the first row (DefaultRowHeight property not available in this version)
                sheet.Cells.SetRowHeight(0, 20);

                // Define output file path
                string outputPath = "output.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
