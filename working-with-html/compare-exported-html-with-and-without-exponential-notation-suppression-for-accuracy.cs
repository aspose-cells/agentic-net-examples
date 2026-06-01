using System;
using System.IO;
using Aspose.Cells;

class CompareHtmlExponentialNotation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Value that would normally be displayed in exponential notation
            double largeNumber = 12345678901234567890.0;
            worksheet.Cells["A1"].PutValue(largeNumber);

            // Apply scientific number format to force exponential notation
            Style sciStyle = worksheet.Cells["A1"].GetStyle();
            sciStyle.Custom = "0.00E+00";
            worksheet.Cells["A1"].SetStyle(sciStyle);

            // Save the workbook (optional verification)
            string outputPath = "ExponentialNotation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}