// Title: Convert HTML to Excel and apply scientific notation custom number format to numeric cells using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load HTML file and format all numeric cells as scientific notation | How to apply custom number format 0.00E+00 to cells after converting HTML to XLSX with Aspose.Cells | Aspose.Cells example for converting HTML to .xlsx and setting scientific notation style programmatically | Load HTML into a workbook with HtmlLoadOptions and set scientific notation for numeric data in C# | Convert HTML table to Excel and automatically format numbers in exponential form using Aspose.Cells
// Tags: html to xlsx conversion using Aspose.Cells | set custom scientific notation format in Aspose.Cells | apply style to numeric cells programmatically C# | load HTML with HtmlLoadOptions Aspose.Cells | iterate cells and set number format C#

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input HTML file, loads it into a Workbook via HtmlLoadOptions, iterates over every used cell in the first worksheet, applies the custom number format "0.00E+00" to each numeric cell, and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.xlsx";

            // Verify that the input HTML file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a new workbook using HtmlLoadOptions
            var loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet (the HTML is loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range of the sheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through all used cells and apply scientific notation to numeric cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        Style style = cell.GetStyle();
                        style.Custom = "0.00E+00"; // Custom format for scientific notation
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
