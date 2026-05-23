using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with diagonal stripe background pattern
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.DiagonalStripe;   // set pattern type
            style.ForegroundColor = Color.LightBlue;         // color of the stripes
            style.BackgroundColor = Color.DarkBlue;          // background color behind the stripes

            // Create the range V1:V10
            AsposeRange range = worksheet.Cells.CreateRange("V1", "V10");

            // Apply the style to the entire range
            range.SetStyle(style);

            // Save the workbook
            string outputPath = "DiagonalStripeStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}