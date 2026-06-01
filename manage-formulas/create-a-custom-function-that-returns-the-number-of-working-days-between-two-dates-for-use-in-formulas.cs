using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // Custom calculation engine that implements the WORKINGDAYS function.
    // The function expects two parameters (start date, end date) and returns
    // the number of working days (Monday‑Friday) between them, inclusive.
    public class WorkingDaysEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // Process only the custom function.
            if (!string.Equals(data.FunctionName, "WORKINGDAYS", StringComparison.OrdinalIgnoreCase))
                return; // Let the default engine handle other functions.

            // Validate parameter count.
            if (data.ParamCount != 2)
            {
                // Return #VALUE! error for incorrect argument count.
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // Retrieve parameters.
            object startObj = data.GetParamValue(0);
            object endObj   = data.GetParamValue(1);

            // Convert to DateTime.
            DateTime startDate, endDate;
            try
            {
                startDate = ConvertToDate(startObj);
                endDate   = ConvertToDate(endObj);
            }
            catch
            {
                // Return #VALUE! error if conversion fails.
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // Ensure startDate <= endDate.
            if (endDate < startDate)
            {
                var tmp = startDate;
                startDate = endDate;
                endDate   = tmp;
            }

            // Count working days (Monday‑Friday) inclusive.
            int workingDays = 0;
            for (DateTime d = startDate; d <= endDate; d = d.AddDays(1))
            {
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday)
                    workingDays++;
            }

            data.CalculatedValue = workingDays;
        }

        // Helper to convert a cell value to DateTime.
        private DateTime ConvertToDate(object value)
        {
            if (value is DateTime dt)
                return dt;

            // Excel numeric dates are stored as double (OADate).
            if (value is double d)
                return DateTime.FromOADate(d);

            // Attempt to parse string representation.
            if (value is string s && DateTime.TryParse(s, out DateTime parsed))
                return parsed;

            throw new InvalidCastException("Unable to convert parameter to DateTime.");
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Input start and end dates.
                cells["A1"].PutValue(new DateTime(2023, 7, 24)); // Monday
                cells["B1"].PutValue(new DateTime(2023, 7, 31)); // Next Monday

                // Use the custom function in a formula.
                cells["C1"].Formula = "=WORKINGDAYS(A1,B1)";

                // Set calculation options to use our custom engine.
                CalculationOptions opts = new CalculationOptions
                {
                    CustomEngine = new WorkingDaysEngine()
                };

                // Calculate all formulas in the workbook.
                wb.CalculateFormula(opts);

                // Output the result.
                Console.WriteLine("Working days between A1 and B1: " + cells["C1"].Value);

                // Save the workbook (optional, demonstrates lifecycle usage).
                string outputPath = "WorkingDaysDemo.xlsx";

                // Ensure the directory exists before saving.
                string fullPath = Path.GetFullPath(outputPath);
                string? dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}