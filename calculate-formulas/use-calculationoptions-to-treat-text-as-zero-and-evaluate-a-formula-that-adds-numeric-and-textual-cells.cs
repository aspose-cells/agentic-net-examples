// Title: Calculate a mixed numeric and text sum in Aspose.Cells C# by enabling TreatTextAsZero option
// AI Prompts: Set workbook.Settings.CalcEngineOptions.TreatTextAsZero = true, assign a formula that adds cells containing numbers and text, then call workbook.CalculateFormula() to obtain the sum. | Create cells with numeric values, numeric strings, and non‑numeric strings, enable TreatTextAsZero, evaluate the formula, and retrieve the resulting value programmatically.
// Common Searches: Aspose.Cells C# treat text values as zero when calculating formulas | How to sum cells that contain numbers and text strings using Aspose.Cells | Enable TreatTextAsZero in Aspose.Cells calculation engine .NET | Formula result 30 with numeric and non‑numeric cells Aspose.Cells example | Ignore non‑numeric text in Excel formula evaluation with Aspose.Cells
// Tags: Aspose.Cells CalculationOptions TreatTextAsZero | C# sum cells with mixed data Aspose.Cells | ignore text values in formula Aspose.Cells | evaluate workbook formulas .NET | numeric string handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, places a numeric value, a numeric string, and a non‑numeric string in cells A1‑A3, sets B1 to the formula =A1+A2+A3, optionally enables the TreatTextAsZero option via CalculationOptions, calculates the formula, prints the result (30 when enabled), and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells: numeric, numeric text, and non‑numeric text
            sheet.Cells["A1"].PutValue(10);          // numeric
            sheet.Cells["A2"].PutValue("20");        // numeric stored as text
            sheet.Cells["A3"].PutValue("abc");       // non‑numeric text

            // Set a formula that adds the three cells
            sheet.Cells["B1"].Formula = "=A1+A2+A3";

            // Treat text values as zero during calculation (available in newer versions)
            // If the property is unavailable, this line can be omitted.
            // workbook.Settings.CalcEngineOptions.TreatTextAsZero = true;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated result (expected: 30 if TreatTextAsZero is true)
            Console.WriteLine("Result: " + sheet.Cells["B1"].Value);

            // Save the workbook (optional)
            string outputPath = "Result.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
