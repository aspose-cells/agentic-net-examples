using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsManualCalcDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet and add a second one
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "DataSheet";
                Worksheet sheet2 = workbook.Worksheets.Add("SummarySheet");

                // -------------------------------------------------
                // Disable automatic calculation for the whole workbook
                // (FormulaSettings.CalculationMode = Manual)
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // -------------------------------------------------
                // Insert sample data and formulas in the first sheet
                // -------------------------------------------------
                Cells cells1 = sheet1.Cells;
                cells1["A1"].PutValue(10);
                cells1["A2"].PutValue(20);
                cells1["A3"].PutValue(30);
                // Simple sum formula
                cells1["B1"].Formula = "=SUM(A1:A3)";

                // -------------------------------------------------
                // Insert formulas in the second sheet that reference the first sheet
                // -------------------------------------------------
                Cells cells2 = sheet2.Cells;
                // Reference a cell from the first sheet
                cells2["A1"].Formula = "=DataSheet!B1";
                // Another calculation based on the first sheet
                cells2["A2"].Formula = "=DataSheet!A1*2";

                // -------------------------------------------------
                // Manually calculate all formulas in the workbook
                // -------------------------------------------------
                workbook.CalculateFormula();

                // -------------------------------------------------
                // Save the workbook (lifecycle rule: save)
                // -------------------------------------------------
                string outputPath = "ManualCalculationDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Output results to console for verification
                Console.WriteLine("DataSheet B1 (SUM): " + cells1["B1"].Value);
                Console.WriteLine("SummarySheet A1 (link to DataSheet B1): " + cells2["A1"].Value);
                Console.WriteLine("SummarySheet A2 (A1*2): " + cells2["A2"].Value);
            }
            catch (Exception ex)
            {
                // Log or display any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}