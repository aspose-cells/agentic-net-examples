// Title: Extract formula text from a defined cell range and create a formula report worksheet with Aspose.Cells for .NET (C#)
// Common Searches: how to extract all formula strings from a range like A1:C10 using Aspose.Cells in C# | Aspose.Cells C# example for reading formulas from a worksheet range | generate a formula report sheet by collecting formulas from a specific range with Aspose.Cells | C# code to iterate over cells and get Formula property using Aspose.Cells | save extracted formulas to a new worksheet using Aspose.Cells for .NET
// Tags: Aspose.Cells extract formulas from range | C# read cell Formula property | create formula report worksheet Aspose.Cells | iterate worksheet cells IsFormula C# | export formula strings to new workbook Aspose.Cells | bulk formula text extraction .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, defines a cell range (A1:C10) on the first worksheet, iterates through each cell, collects the Formula text of cells that contain formulas, writes those formulas to a newly added worksheet named "FormulaReport", and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the address of the range to analyze
            string rangeAddress = "A1:C10";

            // Create an Aspose.Cells.Range (fully qualified to avoid conflict with System.Range)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

            // Collect formula texts
            List<string> formulaTexts = new List<string>();

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;
                    formulaTexts.Add(formula);
                    Console.WriteLine($"Cell {cell.Name} : {formula}");
                }
            }

            // Write extracted formulas to a new worksheet for reporting
            Worksheet reportSheet = workbook.Worksheets.Add("FormulaReport");
            for (int i = 0; i < formulaTexts.Count; i++)
            {
                reportSheet.Cells[i, 0].PutValue(formulaTexts[i]);
            }

            // Save the workbook with the added report sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
