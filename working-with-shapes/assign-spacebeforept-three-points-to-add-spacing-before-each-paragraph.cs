// Title: How to set every row height to 15 points across all worksheets in an existing Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set row height for all rows in a workbook | change row height to specific points in an existing Excel file with Aspose.Cells | iterate through worksheets and rows to adjust row height using Aspose.Cells .NET | set uniform row height 15 points in each worksheet programmatically | C# Aspose.Cells modify row height after loading workbook
// Tags: set row height Aspose.Cells .NET | uniform row height Excel workbook | iterate worksheets rows Aspose.Cells | adjust row height points programmatically | modify existing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook with Aspose.Cells, loops through every worksheet and each row, assigns a Height of 15 points to every row, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and rows, setting row height to 15 points
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Row row in sheet.Cells.Rows)
                    {
                        // Set row height (in points)
                        row.Height = 15;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
