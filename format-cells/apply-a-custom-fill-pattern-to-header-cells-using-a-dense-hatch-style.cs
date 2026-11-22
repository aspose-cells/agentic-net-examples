// Title: How to apply a solid fill pattern to header row cells in an Excel file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set solid background pattern for header cells in C# | C# code to style first row of Excel worksheet using solid fill | Apply custom cell fill color to Excel header using Aspose.Cells .NET | Create and assign a style with solid pattern to multiple cells Aspose.Cells | Load workbook and set header row background color programmatically in .NET
// Tags: solid fill pattern styling Aspose.Cells | header row cell style C# | Excel worksheet background color Aspose.Cells | apply style to range of cells .NET | create custom cell style Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates a style with a solid fill pattern (black foreground on white background), applies this style to the first row cells (A1‑D1) to format them as a header, and saves the updated file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Create a style with a solid fill pattern (Aspose.Cells does not have DenseHatch)
            var headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;      // solid fill pattern
            headerStyle.ForegroundColor = Color.Black;       // fill color
            headerStyle.BackgroundColor = Color.White;       // background color

            // Apply the style to header cells (first row, columns A to D as an example)
            for (int col = 0; col < 4; col++)
            {
                var cell = sheet.Cells[0, col];
                cell.SetStyle(headerStyle);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
