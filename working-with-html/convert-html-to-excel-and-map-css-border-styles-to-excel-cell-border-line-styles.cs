// Title: Convert an HTML file to Excel and translate CSS border styles to Aspose.Cells cell border line types in C#
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook and saves it as an XLSX file. | Create a case‑insensitive dictionary that maps CSS border style names (e.g., solid, dotted, double) to Aspose.Cells CellBorderType enums, then apply the mapped style to each cell’s borders. | Show how to retrieve a CSS border value for a specific cell, set the top, bottom, left, and right borders using Style.SetBorder, and handle any errors during the conversion.
// Common Searches: c# Aspose.Cells load html and preserve table borders | map CSS border property to Excel cell border using Aspose.Cells | convert html table to xlsx with custom border styles in .NET | how to apply dotted border to all sides of a cell with Aspose.Cells | dictionary for CSS to CellBorderType conversion Aspose.Cells example
// Tags: html to xlsx conversion Aspose.Cells | css border to CellBorderType mapping | set uniform cell borders Aspose.Cells | c# dictionary for border style translation | load html workbook Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// This C# example loads an HTML file into an Aspose.Cells Workbook, builds a case‑insensitive dictionary that converts CSS border style names to CellBorderType values, iterates over all used cells, applies the corresponding line style to the top, bottom, left and right borders, and saves the workbook as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the input HTML file.
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Load HTML content into a new workbook using HtmlLoadOptions.
            Workbook workbook;
            try
            {
                workbook = new Workbook(htmlPath, new HtmlLoadOptions());
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load HTML file: {loadEx.Message}");
                return;
            }

            // Mapping from CSS border style names to Aspose.Cells border line styles.
            var cssToExcelBorder = new Dictionary<string, CellBorderType>(StringComparer.OrdinalIgnoreCase)
            {
                { "none",   CellBorderType.None },
                { "solid",  CellBorderType.Thin },
                { "dotted", CellBorderType.Dotted },
                { "dashed", CellBorderType.Dashed },
                { "double", CellBorderType.Double },
                { "groove", CellBorderType.Medium },
                { "ridge",  CellBorderType.Medium },
                { "inset",  CellBorderType.Thick },
                { "outset", CellBorderType.Thick }
            };

            // Ensure there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("No worksheets were created from the HTML content.");
                return;
            }

            // Reference to the first worksheet where HTML was loaded.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells.
            for (int row = cells.MinRow; row <= cells.MaxRow; row++)
            {
                for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Retrieve CSS border style for the current cell (placeholder implementation).
                    string cssBorderStyle = GetCssBorderStyleForCell(row, col);

                    if (!string.IsNullOrEmpty(cssBorderStyle) &&
                        cssToExcelBorder.TryGetValue(cssBorderStyle, out CellBorderType lineStyle))
                    {
                        try
                        {
                            // Get the current style, modify borders, and reapply.
                            Style style = cell.GetStyle();

                            // Apply the mapped line style to all four borders.
                            style.SetBorder(BorderType.BottomBorder, lineStyle, Color.Black);
                            style.SetBorder(BorderType.TopBorder,    lineStyle, Color.Black);
                            style.SetBorder(BorderType.LeftBorder,   lineStyle, Color.Black);
                            style.SetBorder(BorderType.RightBorder,  lineStyle, Color.Black);

                            cell.SetStyle(style);
                        }
                        catch (Exception styleEx)
                        {
                            Console.WriteLine($"Failed to set border for cell [{row}, {col}]: {styleEx.Message}");
                        }
                    }
                }
            }

            // Save the resulting workbook to an Excel file.
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Conversion completed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save the workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder method to simulate retrieval of CSS border style for a specific cell.
    // In practice, implement HTML parsing logic to extract the style.
    static string GetCssBorderStyleForCell(int row, int column)
    {
        // Example stub: return "solid" for demonstration purposes.
        return "solid";
    }
}
