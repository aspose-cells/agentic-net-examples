// Title: Apply a custom Arial style with light‑yellow background to cells A1:C10 in an XLSX workbook and export it as PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing XLSX file, creates a Style with Arial 12‑pt font and a light‑yellow solid fill, applies it to the range A1:C10, and saves the workbook as a PDF using Aspose.Cells. | Show how to use StyleFlag to apply all formatting attributes to a cell range before converting the styled worksheet to PDF in a .NET application.
// Common Searches: aspnet load xlsx apply custom cell style and convert to pdf with aspose.cells | c# aspose.cells apply background color to a range and save as pdf | how to use StyleFlag to apply full formatting to cells before exporting to pdf in aspose.cells | export styled worksheet to pdf preserving cell formatting using aspose.cells .net
// Tags: apply custom style to cell range Aspose.Cells | export styled worksheet to PDF Aspose.Cells | StyleFlag all attributes Aspose.Cells | create Arial font style Aspose.Cells | convert XLSX to PDF with formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, creates an Arial 12‑pt style with a light‑yellow solid fill, applies it to cells A1:C10, and saves the workbook as output.pdf using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Define the range to which the style will be applied (e.g., A1:C10)
            Worksheet worksheet = workbook.Worksheets[0];
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C10");

            // Apply the style to the entire range
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(customStyle, flag);

            // Save the modified workbook as PDF
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{pdfPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
