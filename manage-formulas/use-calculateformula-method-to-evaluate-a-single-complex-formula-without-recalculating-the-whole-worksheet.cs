// Title: Use Worksheet.CalculateFormula to evaluate a complex array formula in Aspose.Cells for .NET without triggering full sheet recalculation
// Common Searches: Aspose.Cells C# calculate a single formula without recalculating the entire workbook | How to evaluate SUM(IF(...)) array formula using Worksheet.CalculateFormula in .NET | Get result of a complex Excel formula and write it to a specific cell with Aspose.Cells | Example of avoiding full worksheet recalculation when computing a conditional sum in C#
// Tags: Worksheet.CalculateFormula API | array formula evaluation Aspose.Cells | conditional sum calculation .NET | write formula result to cell C# | prevent full workbook recalculation

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates one), populates cells A1:A10, B1, and C1, defines a complex formula "SUM(IF(A1:A10>5, A1:A10, 0)) * (B1 + C1)", evaluates it on the first worksheet using Worksheet.CalculateFormula, writes the computed value to cell D1, and saves the file, all without triggering a full worksheet recalculation.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a sample one if the file is missing.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the formula (A1:A10 values).
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1); // A1:A10
                }

                // B1 and C1 values.
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(3);
            }

            // Complex formula to evaluate.
            string formula = "SUM(IF(A1:A10>5, A1:A10, 0)) * (B1 + C1)";

            // Evaluate the formula using the first worksheet.
            Worksheet firstSheet = workbook.Worksheets[0];
            object result = firstSheet.CalculateFormula(formula);

            Console.WriteLine("Calculated result: " + result);

            // Write the result back to a cell without triggering full recalculation.
            firstSheet.Cells["D1"].PutValue(result);

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
