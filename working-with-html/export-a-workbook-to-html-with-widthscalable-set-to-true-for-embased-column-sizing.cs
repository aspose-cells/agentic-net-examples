// Title: Convert an Excel workbook to responsive HTML with WidthScalable enabled using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML with WidthScalable set to true via Aspose.Cells. | Show how to configure HtmlSaveOptions in Aspose.Cells to produce HTML that uses em‑based column sizing for a responsive layout. | Generate a complete .NET example that converts a workbook to HTML with scalable column widths using Aspose.Cells.
// Common Searches: Aspose.Cells export Excel to HTML with responsive column widths | C# HtmlSaveOptions WidthScalable true example | How to generate HTML from a workbook with em based column sizing using Aspose.Cells | Save Excel file as responsive HTML in .NET | Responsive HTML output from Aspose.Cells workbook conversion
// Tags: Aspose.Cells HtmlSaveOptions WidthScalable | export Excel to responsive HTML C# | convert workbook to HTML with em sizing | C# Aspose.Cells HTML export responsive layout | set WidthScalable true Aspose.Cells | HTMLSaveOptions responsive column width

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an existing Excel file, configures HtmlSaveOptions with WidthScalable = true to enable em‑based column sizing, and saves the workbook as a responsive HTML file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Enable scalable width for responsive layout
            htmlOptions.WidthScalable = true;

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
