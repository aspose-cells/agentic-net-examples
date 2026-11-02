// Title: Replace undefined name formula errors (#NAME?) with #REF! in Excel workbooks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change #NAME? error to #REF! after formula calculation | how to detect formulas with undefined names in an Excel file using Aspose.Cells .NET | iterate through worksheets and replace #NAME? cells with #REF! in Aspose.Cells | Aspose.Cells recalculate formulas and fix #NAME? errors programmatically | C# code to convert #NAME? errors to #REF! in an Excel workbook
// Tags: replace #NAME? with #REF! Aspose.Cells | detect undefined name formulas .NET | iterate worksheets cells Aspose.Cells | formula recalculation error handling Excel .NET | update Excel error values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells for .NET, recalculates all formulas, scans every worksheet and cell, identifies cells whose formula result is the #NAME? error (indicating an undefined name), replaces the value with the #REF! error string, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Recalculate all formulas
                workbook.CalculateFormula();

                // Iterate through each worksheet
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    try
                    {
                        Cells cells = worksheet.Cells;

                        // Iterate through all cells in the worksheet
                        foreach (Cell cell in cells)
                        {
                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                // After calculation, cells with undefined names show "#NAME?" error text
                                string cellText = cell.StringValue;
                                if (cellText == "#NAME?")
                                {
                                    // Replace the error with a #REF! error text
                                    cell.PutValue("#REF!");
                                }
                            }
                        }
                    }
                    catch (Exception wsEx)
                    {
                        Console.WriteLine($"Error processing worksheet '{worksheet.Name}': {wsEx.Message}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
