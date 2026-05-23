using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable the 1904 date system for the workbook (optional but consistent)
            workbook.Settings.Date1904 = true;

            // Put a serial date value into A1.
            // In the 1904 system, serial 1 corresponds to 1904-01-02.
            sheet.Cells["A1"].PutValue(1.0);

            // Set a formula in B1 that adds 30 days to the date in A1.
            sheet.Cells["B1"].Formula = "=A1+30";

            // Create calculation options (no need to set Use1904DateSystem here)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate all formulas in the workbook using the specified options.
            workbook.CalculateFormula(calcOptions);

            // Retrieve the calculated serial number from B1.
            double resultSerial = sheet.Cells["B1"].DoubleValue;

            // Convert the serial number to a DateTime using the 1904 epoch.
            DateTime resultDate = CellsHelper.GetDateTimeFromDouble(resultSerial, true);

            // Output the resulting date.
            Console.WriteLine("Calculated date (1904 system): " + resultDate.ToString("yyyy-MM-dd"));

            // Define output file path
            string outputPath = "1904DateSystem.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}