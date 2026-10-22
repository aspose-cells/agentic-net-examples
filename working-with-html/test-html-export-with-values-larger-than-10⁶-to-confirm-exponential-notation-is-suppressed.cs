// Title: How to export Excel cells with values over one million to HTML without scientific notation using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, inserts numbers larger than 1,000,000, applies a custom number format "0" to each cell, and saves the file as HTML with Aspose.Cells, ensuring the output shows plain decimal values. | Show how to configure HtmlSaveOptions and cell styles in Aspose.Cells to suppress exponential notation when converting a worksheet containing large numeric values to HTML.
// Common Searches: Aspose.Cells C# export worksheet to HTML keep large numbers in decimal format | prevent scientific notation for numbers > 10^6 in HTML output with Aspose.Cells | set custom number format for big integers before saving as HTML using Aspose.Cells .NET
// Tags: HTML export suppress scientific notation Aspose.Cells | custom number format integer cells Aspose.Cells | large numeric values HTML save Aspose.Cells | C# Aspose.Cells prevent exponential notation

using Aspose.Cells;
using System;

// Creates a workbook, writes several numbers greater than 1,000,000, applies the custom format "0" to each cell to force plain decimal representation, and saves the workbook as an HTML file using Aspose.Cells, demonstrating how to suppress scientific notation in the HTML output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet and rename it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "LargeNumbers";

            // Large numeric values (greater than 10⁶)
            double[] values = { 1234567, 9876543210, 1.23e6, 5.6789e7 };

            // Write values to cells and apply a custom number format to suppress scientific notation
            for (int i = 0; i < values.Length; i++)
            {
                var cell = sheet.Cells[i, 0];
                cell.PutValue(values[i]);

                // Force plain decimal representation (no exponential format)
                var style = cell.GetStyle();
                style.Custom = "0"; // integer format without scientific notation
                cell.SetStyle(style);
            }

            // Configure HTML export options
            var htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML (format inferred from file extension)
            workbook.Save("LargeNumbers.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
