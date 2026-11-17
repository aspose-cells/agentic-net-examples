// Title: Update cell B4 in an Excel workbook and export to HTML with a custom default font using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change value of cell B4 and save workbook as HTML with Arial font | How to set a default font when converting an Excel file to HTML with Aspose.Cells in .NET | Export an Excel file to HTML using Aspose.Cells while modifying a specific cell value
// Tags: modify worksheet cell Aspose.Cells C# | HTML export with custom default font Aspose | HtmlSaveOptions DefaultFontName usage | Excel to HTML conversion Aspose.Cells .NET | update cell B4 programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of 'input.xlsx', loads it with Aspose.Cells, changes the text in cell B4, sets HtmlSaveOptions.DefaultFontName to 'Arial' to define the default font for the HTML output, and saves the workbook as 'output.html' with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (you can also use workbook.Worksheets["SheetName"])
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify the text of cell B4 (row 3, column 1 in zero‑based indexing)
            Cell targetCell = worksheet.Cells["B4"];
            targetCell.PutValue("Your new text here");

            // Configure HTML save options with a custom default font
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Arial"; // set the desired default font

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
