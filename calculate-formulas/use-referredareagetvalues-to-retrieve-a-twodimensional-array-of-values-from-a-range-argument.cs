// Title: How to retrieve a two‑dimensional object[,] array from a worksheet range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get values from range A1:C3 as object[,] array | Read Excel range into 2D object array with Aspose.Cells for .NET | C# extract worksheet range data as two dimensional matrix using Aspose.Cells | Convert Excel range to object[,] in C# with Aspose.Cells API | Retrieve cell values from a specified range as object matrix in Aspose.Cells
// Tags: Aspose.Cells range Value to object matrix | C# extract Excel cells into two dimensional object array | Aspose.Cells CreateRange API usage | reading worksheet range as object[,] in .NET | Aspose.Cells workbook load and range iteration

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, accesses a worksheet, creates a range (A1:C3), casts the range's Value property to a two‑dimensional object[,] array, iterates the array to display each cell, and saves the workbook, demonstrating error handling for missing files and worksheets.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the worksheet
                Worksheet sheet = workbook.Worksheets["Sheet1"];
                if (sheet == null)
                {
                    Console.WriteLine("Worksheet 'Sheet1' not found.");
                    return;
                }

                // Create a range object for the specified area (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C3");

                // Retrieve the values as a two‑dimensional object array
                object[,] values = range.Value as object[,];
                if (values == null)
                {
                    Console.WriteLine("No data found in the specified range.");
                    return;
                }

                // Iterate and print the retrieved values
                for (int row = 0; row < values.GetLength(0); row++)
                {
                    for (int col = 0; col < values.GetLength(1); col++)
                    {
                        Console.Write(values[row, col] + "\t");
                    }
                    Console.WriteLine();
                }

                // Save the workbook if any modifications were made (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
