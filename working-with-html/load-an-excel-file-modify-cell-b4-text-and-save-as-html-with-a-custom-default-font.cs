// Title: Load an Excel workbook, update cell B4, and export to HTML with a custom default font using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, changes the value of cell B4, sets HtmlSaveOptions.DefaultFontName to a chosen font, and saves the workbook as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active worksheet and apply a specific default font when converting Excel to HTML in C#. | Provide a C# example that verifies the source Excel file exists, updates a cell, and includes exception handling while saving the workbook as HTML with custom font settings.
// Common Searches: Aspose.Cells C# change cell value and save workbook as HTML with specific default font | How to set DefaultFontName in HtmlSaveOptions when converting Excel to HTML using Aspose.Cells | Export only the active worksheet to HTML with custom font using Aspose.Cells .NET | C# code to handle missing input.xlsx file before converting to HTML with Aspose.Cells | Aspose.Cells HTML conversion custom font Calibri example
// Tags: Aspose.Cells HtmlSaveOptions DefaultFontName | C# modify Excel cell before HTML export | Export active worksheet only Aspose.Cells | Excel to HTML conversion with custom font .NET | Error handling missing input file Aspose.Cells | Aspose.Cells workbook.Save HTML with options

using System;
using System.IO;
using Aspose.Cells;

// The program checks for input.xlsx, loads it with Aspose.Cells, updates cell B4 to "New text for B4", configures HtmlSaveOptions with DefaultFontName="Calibri" and ExportActiveWorksheetOnly=true, then saves the result as output.html while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Modify the text of cell B4 (row index 3, column index 1)
            Cell targetCell = sheet.Cells["B4"];
            targetCell.PutValue("New text for B4");

            // Prepare HTML save options with a custom default font
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Set the default font to be used when a cell does not specify a font
                DefaultFontName = "Calibri", // Change to any font you prefer
                // Export only the active worksheet for a smaller output
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as an HTML file using the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
