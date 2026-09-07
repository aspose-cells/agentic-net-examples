// Title: Convert HTML to PDF with Aspose.Cells in C# while preserving exact table border widths
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, applies thin black borders to every populated cell, and saves the workbook as a PDF. | Show how to use HtmlLoadOptions together with cell style settings in Aspose.Cells to keep table border pixel widths unchanged during HTML‑to‑PDF conversion. | Provide a robust C# example that verifies the source HTML file, iterates over all used rows and columns to set border styles, and exports the result to a PDF with precise border rendering.
// Common Searches: how to keep table border thickness when converting html to pdf using aspose.cells c# | asp.net convert html file to pdf with exact cell border widths aspose.cells | c# load html into workbook and export to pdf preserving border pixel size | aspose.cells htmlloadoptions preserve table borders during pdf export
// Tags: html to pdf conversion with aspose.cells c# | set cell border style before pdf export | preserve table border thickness in pdf output | htmlloadoptions workbook loading asp.net | pixel‑perfect borders aspose.cells pdf

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The sample checks for the existence of an input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, iterates over all populated cells in the first worksheet to apply thin black borders on each side, and then saves the workbook as a PDF, ensuring the table borders retain their exact pixel widths.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the HTML file into a workbook
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(inputPath, loadOptions);

            // Apply borders to every cell in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Style style = sheet.Cells[row, col].GetStyle();

                    // Set thin black borders on all sides
                    style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                    style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                    style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                    style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

                    sheet.Cells[row, col].SetStyle(style);
                }
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
