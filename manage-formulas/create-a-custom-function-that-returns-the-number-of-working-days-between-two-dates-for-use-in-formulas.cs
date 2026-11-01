// Title: Calculate working days between two dates in an Aspose.Cells .NET workbook using the NETWORKDAYS formula
// Common Searches: Aspose.Cells C# example for calculating business days with NETWORKDAYS | How to use NETWORKDAYS function in Aspose.Cells to count weekdays | C# code to compute working days between dates in an Excel file using Aspose.Cells | Aspose.Cells workbook formula for number of weekdays between two dates | Calculate net working days in .NET Excel workbook with Aspose.Cells
// Tags: Aspose.Cells formula execution workflow | C# set cell values and trigger calculation | Excel workbook business day count using Aspose | persist workbook after NETWORKDAYS evaluation | date cell initialization for formula in Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    // The sample creates a new workbook, writes start and end dates to cells A1 and B1, applies the NETWORKDAYS formula in C1 to compute the number of working days, forces formula evaluation, saves the workbook as WorkingDaysDemo.xlsx, and outputs a confirmation message.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample dates: A1 = start date, B1 = end date.
                sheet.Cells["A1"].PutValue(new DateTime(2023, 4, 3)); // Monday
                sheet.Cells["B1"].PutValue(new DateTime(2023, 4, 14)); // Friday

                // Use the built‑in NETWORKDAYS function to calculate working days.
                // C1 will display the result.
                sheet.Cells["C1"].Formula = "=NETWORKDAYS(A1, B1)";

                // Calculate formulas so that C1 gets the computed value.
                workbook.CalculateFormula();

                // Save the workbook.
                string outputPath = "WorkingDaysDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
