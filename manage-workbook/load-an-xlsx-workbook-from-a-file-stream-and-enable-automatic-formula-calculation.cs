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
            string inputPath = "input.xlsx";

            // Open the file as a read‑only stream
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Create load options (default settings)
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook from the stream using the constructor that accepts a Stream and LoadOptions
                Workbook workbook = new Workbook(stream, loadOptions);

                // Enable automatic formula calculation by setting the calculation mode to Automatic
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // (Optional) Force calculation now if you need the results immediately
                // workbook.CalculateFormula();

                // Example: read a calculated cell value after enabling automatic calculation
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("Value of A1 after loading: " + sheet.Cells["A1"].Value);
            }
        }
    }
}