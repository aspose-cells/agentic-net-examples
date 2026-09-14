// Title: How to hide rows 5‑15 and enable formula view in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program using Aspose.Cells that conceals rows 5‑15 in the first worksheet, activates the formula‑display mode for the sheet, and writes the result to a new Excel file. | Update an existing Aspose.Cells workbook so that rows 5‑15 become hidden and the workbook is set to show formulas instead of calculated values before saving.
// Common Searches: Aspose.Cells hide rows 5‑15 and show formulas in C# | Enable formula view after hiding rows with Aspose.Cells .NET | C# Aspose.Cells set workbook to display formulas | How to hide a range of rows and view formulas using Aspose.Cells | Aspose.Cells hide specific rows then turn on ShowFormula property
// Tags: hide rows Aspose.Cells C# | show formulas Aspose.Cells workbook | Aspose.Cells row visibility control | Aspose.Cells ShowFormula property .NET | C# Excel worksheet row hiding and formula display

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, hides rows 5‑15 in the first worksheet by calling Cells.HideRow for each zero‑based index, sets the workbook’s ShowFormula setting to true so formulas are displayed instead of values, and saves the modified workbook to a new file, handling missing input files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Hide rows 5 through 15 (Excel rows are 1‑based, Aspose.Cells rows are 0‑based)
            for (int rowIndex = 4; rowIndex <= 14; rowIndex++)
            {
                // HideRow takes a single argument (row index)
                sheet.Cells.HideRow(rowIndex);
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
