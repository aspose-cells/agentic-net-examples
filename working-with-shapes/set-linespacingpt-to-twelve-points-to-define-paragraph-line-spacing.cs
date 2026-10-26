// Title: How to set a 12‑point row height in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set specific row height in points | set Excel row height to 12 points using Aspose.Cells .NET | exact row height in points with Aspose.Cells workbook | C# Aspose.Cells row height precision example | define first row height 12pt in generated Excel file Aspose.Cells
// Tags: row height points Aspose.Cells | Excel row height precision C# | Aspose.Cells set exact row height | worksheet row height in points | C# Aspose.Cells row height example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, writes text to cell A1, sets the first row's height to exactly 12 points, and saves the file as LineSpacingExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add some text to a cell
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Line spacing example");

                // Set the row height to exactly 12 points
                // Row index is zero‑based; 0 refers to the first row
                sheet.Cells.SetRowHeight(0, 12);

                // Define output file path
                string outputPath = "LineSpacingExample.xlsx";

                // Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
