// Title: Identify and Highlight #DIV/0! Formula Errors in an Excel Workbook Using Aspose.Cells for .NET (C#)
// Common Searches: C# code to locate #DIV/0! cells in an Excel file with Aspose.Cells | how to highlight division by zero errors in Excel using Aspose.Cells .NET | collect addresses of formula errors across all worksheets with Aspose.Cells | apply background color to error cells in a workbook using Aspose.Cells C#
// Tags: detect division by zero errors Aspose.Cells | apply error cell background Aspose.Cells | iterate worksheets collect formula errors | use IsErrorValue to find #DIV/0! Aspose.Cells | save workbook after error highlighting C# | store cell addresses of formula errors

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsFormulaErrorCheck
{
    // The example loads an XLSX workbook with Aspose.Cells, scans every worksheet's used range, detects cells where IsErrorValue is true and the string equals "#DIV/0!", records each cell's A1 address with its sheet name, prints the list, applies a LightCoral solid background to those cells, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // List to store addresses of cells that contain #DIV/0! errors
                List<string> divZeroCells = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Iterate through each cell in the used range
                    foreach (Cell cell in usedRange)
                    {
                        // Check if the cell contains an error value and specifically #DIV/0!
                        if (cell.IsErrorValue && cell.StringValue == "#DIV/0!")
                        {
                            // Store the cell address in A1 style with sheet name
                            string address = $"{sheet.Name}!{cell.Name}";
                            divZeroCells.Add(address);
                        }
                    }
                }

                // Output the collected addresses
                Console.WriteLine("Cells with #DIV/0! errors:");
                foreach (string addr in divZeroCells)
                {
                    Console.WriteLine(addr);
                }

                // Highlight the error cells (e.g., set background color)
                foreach (string addr in divZeroCells)
                {
                    // Split sheet name and cell address
                    int exclPos = addr.IndexOf('!');
                    string sheetName = addr.Substring(0, exclPos);
                    string cellRef = addr.Substring(exclPos + 1);

                    Worksheet sheet = workbook.Worksheets[sheetName];
                    Cell cell = sheet.Cells[cellRef];

                    // Apply a light red background to highlight the error cell
                    Style style = cell.GetStyle();
                    style.ForegroundColor = Color.LightCoral;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
