using System;
using System.Drawing;
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

            // Define the target range N5:N15 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("N5:N15");

            // Create a style and set thin black borders on all four sides
            Style style = workbook.CreateStyle();
            style.SetBorder(BorderType.SideBorders, CellBorderType.Thin, Color.Black);

            // Apply the style to the range
            range.SetStyle(style);

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}