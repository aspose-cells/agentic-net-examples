// Title: Prevent scientific notation for large numeric values when exporting an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through every cell in a workbook and sets a plain integer format ("0") for numeric cells before saving the file as HTML with Aspose.Cells. | Show how to configure the HTML export options in Aspose.Cells so that the generated HTML respects the integer format and does not display numbers in exponential form. | Demonstrate a technique to identify numeric cells and apply an integer display style to avoid scientific notation during Excel‑to‑HTML conversion in .NET.
// Common Searches: Aspose.Cells C# export Excel to HTML without scientific notation for large numbers | How to keep large numeric values from showing as 1E+09 in HTML output using Aspose.Cells | Set integer format for numeric cells before saving workbook as HTML in .NET | Prevent exponential notation when converting Excel to HTML with Aspose.Cells library | C# Aspose.Cells HTML export preserve integer display for big numbers
// Tags: Aspose.Cells custom number format for HTML export | C# suppress scientific notation in HTML output | Excel HTML export large numbers Aspose.Cells | HtmlSaveOptions numeric formatting Aspose.Cells | apply integer format to numeric cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over all worksheets and cells, applies a plain integer format ("0") to every numeric cell to stop scientific notation, and then saves the workbook as HTML using HtmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the source workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their used cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Apply custom number format only to numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            Style style = cell.GetStyle();
                            // Use a plain integer format to suppress scientific notation
                            style.Custom = "0";
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Configure HTML export options (default settings already respect cell formats)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
