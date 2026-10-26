// Title: Populate an Excel worksheet range from a two‑dimensional object[,] array using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells write a 2D object array to a specific Excel range | populate range A1:B2 dynamically with object[,] using Aspose.Cells | loop through object[,] and put values into worksheet cells Aspose.Cells | save workbook after filling cells from a two‑dimensional array in C#
// Tags: populate worksheet range from object array Aspose.Cells | write two-dimensional array to Excel cells C# | Aspose.Cells PutValue loop example | dynamic range filling Aspose.Cells .NET | object[,] to Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, defines a 2×2 object[,] array, iterates over its rows and columns, writes each element to the corresponding worksheet cell with PutValue, ensures the output directory exists, and saves the file as DynamicArrayResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a 2x2 array of values to populate the range A1:B2
            object[,] data = new object[,]
            {
                { "Item1", "Item2" },
                { "Item3", "Item4" }
            };

            // Manually write the array values to the worksheet (avoids ImportArray overload issues)
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    sheet.Cells[row, col].PutValue(data[row, col]);
                }
            }

            // Define the output file path
            string outputPath = "DynamicArrayResult.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
