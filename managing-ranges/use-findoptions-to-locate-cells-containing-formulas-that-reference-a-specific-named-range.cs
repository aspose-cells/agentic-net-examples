// Title: Locate formulas that reference a specific named range in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells FindOptions to locate formulas that contain a named range | how to search Excel formulas for a particular named range with Aspose.Cells | retrieve cells referencing MyRange using Aspose.Cells FindOptions in .NET | scan workbook for formulas referencing a specific named range using Aspose.Cells API | list cells that use a given named range in their formula with Aspose.Cells C#
// Tags: FindOptions formula search Aspose.Cells | named range reference detection Aspose.Cells | locate cells referencing named range Excel | scan workbook formulas with Aspose.Cells | C# Excel formula lookup using FindOptions

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and cell, checks each formula for the specified named range, prints the worksheet name, cell address and formula, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string targetNamedRange = "MyRange";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Scan each worksheet for formulas containing the named range
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf(targetNamedRange, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cell.Name}, Formula: {formula}");
                        }
                    }
                }
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
