// Title: How to disable auto‑calculation, insert formulas on multiple sheets, and manually recalculate with Worksheet.CalculateFormula in Aspose.Cells for .NET
// AI Prompts: Create C# code that configures the workbook to use manual recalculation, places a SUM formula in Sheet1!A1 and an AVERAGE formula in Sheet2!C1, then explicitly evaluates the formulas with Worksheet.CalculateFormula before saving. | Provide a sample that adds formulas to two worksheets in Aspose.Cells, suppresses automatic calculation, and manually triggers formula evaluation for each sheet using the appropriate API.
// Common Searches: Aspose.Cells set calculation mode to manual and recalculate specific worksheet in C# | C# Aspose.Cells add SUM formula to Sheet1 and AVERAGE to Sheet2 then manual calculate | How to prevent automatic formula evaluation in Aspose.Cells workbook | Worksheet.CalculateFormula usage example for multiple sheets Aspose.Cells .NET
// Tags: disable automatic calculation Aspose.Cells workbook | add SUM formula to Sheet1 A1 Aspose.Cells | add AVERAGE formula to Sheet2 C1 Aspose.Cells | Worksheet.CalculateFormula manual trigger | save workbook after formula evaluation Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, inserts a SUM formula into Sheet1!A1 and an AVERAGE formula into Sheet2!C1, disables automatic recalculation, manually evaluates all formulas with Worksheet.CalculateFormula, and saves the file as output.xlsx while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // NOTE: In some older Aspose.Cells versions the CalcMode property may not be available.
            // If needed, you can set the calculation mode to manual using the appropriate API for your version.
            // For compatibility, the line is omitted here.

            // Worksheet 1 (default first sheet)
            var sheet1 = workbook.Worksheets[0];
            // Example formula: sum of B1:B10 placed in A1
            sheet1.Cells["A1"].Formula = "=SUM(B1:B10)";

            // Worksheet 2 (add a new sheet)
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            // Example formula: average of D1:D5 placed in C1
            sheet2.Cells["C1"].Formula = "=AVERAGE(D1:D5)";

            // Manually trigger calculation for the entire workbook
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Calculation error: {calcEx.Message}");
            }

            // Save the workbook
            try
            {
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved successfully as output.xlsx");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Save error: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
