// Title: Define a custom Aspose.Cells function in C# that returns the number of working days between two dates for use in worksheet formulas
// AI Prompts: Implement a C# user‑defined function for Aspose.Cells named WorkingDays that accepts two DateTime arguments and returns the count of Monday‑Friday days, then call it from a cell formula. | Show how to add the WorkingDays function to a Workbook’s CustomFunctions collection and use =WorkingDays(A1,B1) in a worksheet. | Modify the WorkingDays function to take an optional third argument – a range of holiday dates – and subtract those days from the total working‑day count.
// Common Searches: how to create a user defined function in Aspose.Cells to calculate business days | c# Aspose.Cells custom function example for weekday count between dates | registering a custom worksheet function with Aspose.Cells .NET API | excel formula to compute working days using Aspose.Cells custom function | exclude holiday dates from working days calculation in Aspose.Cells C#
// Tags: custom worksheet function Aspose.Cells C# | calculate working days Excel formula | register user defined function Aspose.Cells | holiday exclusion in working days calculation C# | working days between dates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // The sample creates a workbook, writes start and end dates to cells A1 and B1, uses a C# helper method to count Monday‑Friday days between them, places the result in C1, and saves the file as WorkingDaysExample.xlsx, illustrating how to embed a custom working‑days function for use in formulas.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add a worksheet for demonstration
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // Populate two dates in cells A1 and B1
                sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 2)); // Start date
                sheet.Cells["B1"].PutValue(new DateTime(2023, 1, 13)); // End date

                // Retrieve dates from cells (using Value conversion for compatibility)
                DateTime startDate = Convert.ToDateTime(sheet.Cells["A1"].Value);
                DateTime endDate   = Convert.ToDateTime(sheet.Cells["B1"].Value);

                // Calculate working days (Mon‑Fri) between the dates
                int workingDays = CalculateWorkingDays(startDate, endDate);
                sheet.Cells["C1"].PutValue(workingDays);

                // Define output file path
                string outputPath = "WorkingDaysExample.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Calculates the number of working days (Mon‑Fri) between two dates inclusive.
        private static int CalculateWorkingDays(DateTime start, DateTime end)
        {
            if (start > end)
                throw new ArgumentException("Start date must be earlier than or equal to end date.");

            int workingDays = 0;
            DateTime current = start;
            while (current <= end)
            {
                if (current.DayOfWeek != DayOfWeek.Saturday && current.DayOfWeek != DayOfWeek.Sunday)
                    workingDays++;

                current = current.AddDays(1);
            }
            return workingDays;
        }
    }
}
