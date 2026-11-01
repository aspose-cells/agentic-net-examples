// Title: Convert HTML tables to Excel and format all numbers in scientific notation using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells HtmlLoadOptions, iterates through every cell, converts any numeric or numeric‑string values to double, applies the custom number format "0.00E+00", and saves the workbook as an .xlsx file. | Update an existing Aspose.Cells HTML‑to‑Excel conversion so that string cells containing numbers are parsed to doubles and styled with scientific notation. | Add robust error handling to a C# Aspose.Cells program that verifies the source HTML file exists, catches load failures, and logs clear messages before exiting.
// Common Searches: how to load an html file into Aspose.Cells workbook and preserve numeric values | Aspose.Cells C# convert html table to xlsx with scientific notation for numbers | apply custom number format 0.00E+00 to cells after importing html using Aspose.Cells | detect numeric strings in imported html and format as exponential notation in Aspose.Cells .NET
// Tags: html to xlsx conversion Aspose.Cells | apply scientific notation format Aspose.Cells | custom number format 0.00E+00 C# | convert string to numeric Aspose.Cells | load html with HtmlLoadOptions Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program loads an HTML file into an Aspose.Cells Workbook, walks through each worksheet's cells, converts numeric and parseable string values to doubles, applies the custom scientific notation format "0.00E+00" to those cells, and saves the result as an XLSX workbook.
class Program
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the HTML file into a new workbook using HtmlLoadOptions
            Workbook workbook;
            try
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load HTML file: {loadEx.Message}");
                return;
            }

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Numeric cells: apply scientific format
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            Style style = cell.GetStyle();
                            style.Custom = "0.00E+00";
                            cell.SetStyle(style);
                        }
                        // String cells that can be parsed as numbers: convert and format
                        else if (cell.Type == CellValueType.IsString)
                        {
                            if (double.TryParse(cell.StringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue))
                            {
                                cell.PutValue(numericValue);
                                Style style = cell.GetStyle();
                                style.Custom = "0.00E+00";
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
