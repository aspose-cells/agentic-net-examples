using System;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // ----- Data import section (sample data) -----
            // Add some data to the first worksheet
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Sheet1";
            ws1.Cells["A1"].PutValue(10);
            ws1.Cells["A2"].PutValue(20);
            ws1.Cells["B1"].Formula = "=A1+A2";

            // Add a second worksheet and import more data
            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            ws2.Cells["A1"].PutValue(5);
            ws2.Cells["A2"].PutValue(15);
            ws2.Cells["B1"].Formula = "=SUM(A1:A2)";

            // ----- End of data import -----

            // Iterate over all worksheets and set calculation mode to Automatic
            // (CalculationMode is a workbook‑level setting, but we follow the requirement)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Set the calculation mode to Automatic for the workbook
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                // Optionally, you could recalculate formulas now
                // workbook.CalculateFormula();
            }

            // Save the workbook to verify the setting is persisted
            workbook.Save("Output_AutomaticCalculation.xlsx", SaveFormat.Xlsx);
        }
    }
}