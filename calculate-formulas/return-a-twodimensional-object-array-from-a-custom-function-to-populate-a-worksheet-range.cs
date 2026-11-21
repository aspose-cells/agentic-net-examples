// Title: Write a two‑dimensional object[,] array to an Excel range with Aspose.Cells in C#
// Common Searches: how to assign a 2D object array to a worksheet range using Aspose.Cells C# | populate Excel cells from a custom function returning object[,] in C# | create a range that matches array dimensions and set its value with Aspose.Cells | save workbook after filling range with object[,] data in C# Aspose.Cells
// Tags: Aspose.Cells set range value from object[,] array | C# create matching range dimensions for 2D array | populate Excel worksheet range with object array Aspose.Cells | write two dimensional data to XLSX using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, obtains a two‑dimensional object[,] from a custom method, builds a range whose size matches the array starting at A1, assigns the array to the range's Value property, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve a two‑dimensional object array from a custom function
            object[,] data = GetSampleData();

            // Determine the size of the array
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            // Create a range that matches the array dimensions (starting at A1)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, rows, cols);

            // Populate the range with the array values
            range.Value = data;

            // Ensure the output directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Custom function that returns a two‑dimensional object array
    static object[,] GetSampleData()
    {
        // Example: 3 rows x 3 columns
        object[,] arr = new object[3, 3];
        arr[0, 0] = "Name";
        arr[0, 1] = "Age";
        arr[0, 2] = "Score";

        arr[1, 0] = "Alice";
        arr[1, 1] = 30;
        arr[1, 2] = 85.5;

        arr[2, 0] = "Bob";
        arr[2, 1] = 28;
        arr[2, 2] = 92.0;

        return arr;
    }
}
