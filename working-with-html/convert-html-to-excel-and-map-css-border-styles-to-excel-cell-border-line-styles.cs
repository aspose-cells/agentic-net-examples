// Title: Convert HTML to XLSX with Aspose.Cells while translating CSS border styles to Excel cell borders in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to load an HTML file via HtmlLoadOptions, map CSS border style names to CellBorderType values, apply the mapped borders to each cell, and save the workbook as an XLSX file. | Create a method that parses CSS border definitions from an HTML table and returns the corresponding Aspose.Cells CellBorderType and Color, then integrate it into the HTML‑to‑Excel conversion workflow. | Extend the example to support CSS border widths by selecting the appropriate CellBorderType (Thin, Medium, Thick) based on the numeric width value extracted from the HTML.
// Common Searches: preserve HTML table borders during Aspose.Cells HTML to XLSX conversion C# | translate CSS border definitions to Excel cell borders using Aspose.Cells | customize cell border colors from HTML CSS when exporting to XLSX in C# | load HTML with HtmlLoadOptions and set Excel cell border line styles Aspose.Cells | map CSS border widths to Aspose.Cells CellBorderType values C#
// Tags: HTML to XLSX conversion Aspose.Cells | CSS border style mapping CellBorderType | apply Excel cell borders from HTML CSS | HtmlLoadOptions workbook loading C# | custom cell border line styles Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample reads an HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, defines a dictionary that maps CSS border style names (e.g., solid, double, dashed) to CellBorderType values, iterates through every worksheet and cell, obtains a placeholder CSS border style for each cell, applies the corresponding Excel border type to all four sides with a black color, ensures the output directory exists, and saves the result as an XLSX workbook while handling errors.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.xlsx";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Read the HTML source
            string html = File.ReadAllText(inputPath, Encoding.UTF8);

            Workbook workbook = null;

            // Load the HTML content into the workbook using a memory stream and HtmlLoadOptions
            try
            {
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
                {
                    HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                    workbook = new Workbook(ms, loadOptions);
                }
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load HTML into workbook: {loadEx.Message}");
                return;
            }

            // Mapping from CSS border style names to Aspose.Cells border line styles
            var cssToExcelBorder = new Dictionary<string, CellBorderType>(StringComparer.OrdinalIgnoreCase)
            {
                { "none",   CellBorderType.None },
                { "solid",  CellBorderType.Thin },
                { "double", CellBorderType.Double },
                { "dashed", CellBorderType.Dashed },
                { "dotted", CellBorderType.Dotted },
                { "thick",  CellBorderType.Thick }
            };

            // Iterate through all worksheets and cells to apply the mapped border styles
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    // Retrieve the CSS border style for the current cell.
                    // Implement GetCssBorderStyle to parse the original HTML/CSS and return a style name (e.g., "solid").
                    string cssBorderStyle = GetCssBorderStyle(cell.Name);

                    if (cssToExcelBorder.TryGetValue(cssBorderStyle, out CellBorderType excelBorder))
                    {
                        // Get the current style of the cell
                        Style style = cell.GetStyle();

                        // Apply the mapped border style to all four sides
                        style.SetBorder(BorderType.TopBorder,    excelBorder, Color.Black);
                        style.SetBorder(BorderType.BottomBorder, excelBorder, Color.Black);
                        style.SetBorder(BorderType.LeftBorder,   excelBorder, Color.Black);
                        style.SetBorder(BorderType.RightBorder,  excelBorder, Color.Black);

                        // Assign the updated style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder method – replace with actual logic to extract CSS border style for a given cell address
    static string GetCssBorderStyle(string cellAddress)
    {
        // Example implementation: return "solid" for all cells.
        // In a real scenario, parse the HTML/CSS to determine the border style per cell.
        return "solid";
    }
}
