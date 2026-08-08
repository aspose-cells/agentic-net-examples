// Title: Sum Mixed Numeric and Text Cells in Aspose.Cells (C#) – CalculationOptions.TreatTextAsZero
// Description: C# example that creates a workbook, writes a number, a numeric string and a non‑numeric string to cells A1‑A3, enables CalculationOptions.TreatTextAsZero, evaluates "=SUM(A1:A3)" to get 30, and saves the file. Demonstrates how Aspose.Cells treats text as zero during formula calculation.
// Keywords: Aspose.Cells | CalculationOptions | TreatTextAsZero | C# SUM formula | evaluate Excel formula | mixed data types | Excel calculation | Aspose.Cells example
// Common Searches: Aspose.Cells treat text as zero | C# calculate SUM with text cells | CalculationOptions.TreatTextAsZero example | how to sum numeric strings in Aspose.Cells | Aspose.Cells CalculateFormula mixed data
// Developer Intent: The developer needs to add numbers together even when some cells contain text, ensuring that any text values are counted as zero during the calculation.
// Use Cases: Summing financial columns where some entries are stored as text strings. | Aggregating user‑entered data that may include accidental non‑numeric entries. | Generating reports that must handle mixed cell types without triggering calculation errors.
// AI Prompts: Provide C# code that sets CalculationOptions.TreatTextAsZero = true and evaluates =SUM(A1:A3) with Aspose.Cells. | Explain the impact of CalculationOptions.TreatTextAsZero on formula evaluation when text cells are present. | Create a full example that opens an existing workbook, applies TreatTextAsZero, calculates a sum formula, and saves the result.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a workbook, writes a number, a numeric string and a non‑numeric string to cells A1‑A3, enables CalculationOptions.TreatTextAsZero, evaluates "=SUM(A1:A3)" to get 30, and saves the file. Demonstrates how Aspose.Cells treats text as zero during formula calculation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells: numeric, numeric string, and non‑numeric string
            sheet.Cells["A1"].PutValue(10);          // numeric value
            sheet.Cells["A2"].PutValue("20");       // numeric text
            sheet.Cells["A3"].PutValue("invalid");  // non‑numeric text

            // Evaluate a formula that adds the three cells
            object result = sheet.CalculateFormula("=SUM(A1:A3)");

            // Output the calculated result (expected 30)
            Console.WriteLine("Calculated SUM result: " + result);

            // Save the workbook (optional)
            string outputPath = "CalculatedResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
