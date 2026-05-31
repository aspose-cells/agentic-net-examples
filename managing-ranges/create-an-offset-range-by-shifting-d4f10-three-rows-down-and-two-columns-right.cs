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
            Cells cells = worksheet.Cells;

            // Create the original range D4:F10
            Aspose.Cells.Range originalRange = cells.CreateRange("D4", "F10");

            // Shift the range 3 rows down and 2 columns to the right
            Aspose.Cells.Range offsetRange = originalRange.GetOffset(3, 2);

            // Display the addresses of both ranges
            Console.WriteLine("Original Range Address: " + originalRange.Address);
            Console.WriteLine("Offset Range Address: " + offsetRange.Address);

            // Define output file path
            string outputPath = "OffsetRangeDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}