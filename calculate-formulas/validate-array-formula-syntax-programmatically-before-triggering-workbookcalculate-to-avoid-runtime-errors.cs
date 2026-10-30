// Title: How to programmatically verify Excel array formula syntax with Aspose.Cells for .NET before calling Workbook.CalculateFormula
// Common Searches: C# check Excel array formula syntax with Aspose.Cells before calculation | Aspose.Cells .NET detect invalid formulas to prevent Workbook.CalculateFormula errors | Validate Excel formulas programmatically using Aspose.Cells and handle calculation exceptions | How to catch formula calculation failures in Aspose.Cells workbook in C# | Pre‑validate workbook formulas with Aspose.Cells to avoid runtime errors
// Tags: array formula validation Aspose.Cells | Workbook.CalculateFormula exception handling | pre‑calculation formula check .NET | Excel formula error detection Aspose.Cells | programmatic formula validation C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example demonstrates loading or creating a workbook, attempting to calculate all formulas with Workbook.CalculateFormula, catching any exceptions that indicate invalid array formula syntax, logging the error details, and finally saving the workbook. It shows how to programmatically validate formula syntax and handle calculation errors using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook or create a new one if the file does not exist
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Attempt to calculate all formulas; if an error occurs, capture details
            bool calculationSucceeded = true;
            StringBuilder calculationErrors = new StringBuilder();

            try
            {
                // Calculate all formulas in the workbook
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                calculationSucceeded = false;
                calculationErrors.AppendLine("Formula calculation failed: " + calcEx.Message);
            }

            if (!calculationSucceeded)
            {
                Console.WriteLine("One or more formulas could not be calculated.");
                Console.WriteLine(calculationErrors.ToString());
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
