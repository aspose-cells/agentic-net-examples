using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadAndCalculate
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string filePath = "input.xlsx";

            // Open the file as a read‑only stream
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream
                Workbook workbook = new Workbook(stream);

                // Enable automatic calculation mode (explicitly set for clarity)
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // Optionally, calculate formulas immediately if needed
                workbook.CalculateFormula();

                // Example: read a calculated cell value
                Console.WriteLine("A1 value after calculation: " + workbook.Worksheets[0].Cells["A1"].Value);
            }
        }
    }
}