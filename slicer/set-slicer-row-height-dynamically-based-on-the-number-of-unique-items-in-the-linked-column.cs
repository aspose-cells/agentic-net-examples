// Title: How to set an Excel slicer’s height dynamically based on unique values in a column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# adjust slicer height according to number of distinct items in a column | set slicer row height dynamically based on unique values with Aspose.Cells .NET | calculate distinct values in Excel column and resize slicer height using Aspose.Cells | C# code to change slicer size after counting unique entries in linked column | Aspose.Cells example for dynamic slicer height based on data count
// Tags: Aspose.Cells set slicer height programmatically | calculate distinct values in worksheet column C# | auto resize Excel slicer based on data count | C# Aspose.Cells slicer resizing example | Excel slicer height scaling with unique items

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer class

namespace AsposeCellsDynamicSlicer
{
    // The sample loads an existing workbook, gathers distinct non‑null values from a specified column, retrieves the first slicer on the sheet, and sets the slicer's Height property to the count of unique items multiplied by a fixed item height, then saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                string inputPath = @"C:\Input\Sample.xlsx";
                string outputPath = @"C:\Output\Sample_With_Dynamic_Slicer.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Column that the slicer is linked to (e.g., column B)
                int linkedColumnIndex = 1; // zero‑based (0 = A, 1 = B, ...)

                // Determine the data range in the linked column
                int firstDataRow = 0;                     // zero‑based (row 1 in Excel)
                int lastDataRow = sheet.Cells.MaxDataRow; // last used row in the sheet

                // Collect unique non‑null values from the linked column
                HashSet<object> uniqueValues = new HashSet<object>();
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    object val = sheet.Cells[row, linkedColumnIndex].Value;
                    if (val != null)
                        uniqueValues.Add(val);
                }

                // Ensure the worksheet contains at least one slicer
                if (sheet.Slicers.Count == 0)
                {
                    Console.WriteLine("No slicer found on the worksheet.");
                    return;
                }

                // Get the first slicer (adjust index if multiple slicers are present)
                Slicer slicer = sheet.Slicers[0];

                // Desired height for each slicer item (in points)
                double itemHeight = 20.0;

                // Adjust the overall slicer height based on the number of unique items
                // Use the Height property (measured in points)
                slicer.Height = uniqueValues.Count * itemHeight;

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
