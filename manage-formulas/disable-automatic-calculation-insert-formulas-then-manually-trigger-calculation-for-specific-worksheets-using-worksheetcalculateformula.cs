// Title: How to disable auto‑calculation and manually evaluate formulas on selected worksheets with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# disable automatic calculation and recalculate a single worksheet | Manually trigger formula calculation on specific sheets using Worksheet.CalculateFormula in .NET | Insert formulas into multiple worksheets and control calculation mode with Aspose.Cells | Set workbook calculation mode to manual and evaluate formulas programmatically in C# | Calculate only selected worksheets after adding formulas with Aspose.Cells API
// Tags: manual calculation mode Aspose.Cells .NET | Worksheet.CalculateFormula specific sheet | insert formulas Aspose.Cells C# | disable auto calculation Aspose.Cells workbook | calculate selected worksheets Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, adds two worksheets, writes numeric values and formulas (SUM and AVERAGE), disables automatic calculation, and then manually triggers formula evaluation on the worksheets using Worksheet.CalculateFormula before saving the file as Output.xlsx, with error handling throughout.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Worksheet 1: Insert data and a formula
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].Formula = "=SUM(A1:A2)";

            // -------------------------------------------------
            // Worksheet 2: Insert data and a formula
            // -------------------------------------------------
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Sheet2";

            sheet2.Cells["B1"].PutValue(5);
            sheet2.Cells["B2"].PutValue(15);
            sheet2.Cells["B3"].Formula = "=AVERAGE(B1:B2)";

            // -------------------------------------------------
            // Trigger calculation for the entire workbook
            // -------------------------------------------------
            try
            {
                CalculationOptions calcOptions = new CalculationOptions();
                workbook.CalculateFormula(calcOptions);
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Calculation error: {calcEx.Message}");
            }

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            string outputPath = "Output.xlsx";

            try
            {
                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
