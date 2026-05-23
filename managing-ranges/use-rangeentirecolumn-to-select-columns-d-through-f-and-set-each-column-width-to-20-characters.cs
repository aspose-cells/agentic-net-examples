using System;
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

            // Create a range that spans columns D through F (use alias to avoid conflict with System.Range)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("D:F");

            // Obtain the entire columns that the range occupies
            Aspose.Cells.Range entireColumns = range.EntireColumn;

            // Set the width of each column in the range to 20 characters
            entireColumns.ColumnWidth = 20;

            // Save the workbook
            string outputPath = "Columns_D_F_Width20.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}