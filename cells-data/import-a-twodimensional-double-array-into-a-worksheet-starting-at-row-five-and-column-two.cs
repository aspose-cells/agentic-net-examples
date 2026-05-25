using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a two‑dimensional object array with sample data (values will be boxed)
            object[,] data = new object[,]
            {
                { 1.23, 4.56, 7.89 },
                { 10.11, 12.13, 14.15 },
                { 16.17, 18.19, 20.21 }
            };

            // Import the array starting at row 5 (zero‑based index 4) and column 2 (zero‑based index 1)
            worksheet.Cells.ImportTwoDimensionArray(data, 4, 1);

            // Define output file path
            string outputPath = "TwoDimensionalArrayImport.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}