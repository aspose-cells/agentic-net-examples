// Title: Calculate working days between two dates with Aspose.Cells NETWORKDAYS formula in C#
// Common Searches: Aspose.Cells C# example for NETWORKDAYS to get business days count | How to use NETWORKDAYS function in Aspose.Cells workbook calculation | C# code to compute workdays between dates in Excel using Aspose.Cells | Calculate business days in .NET Excel file with Aspose.Cells formula | Retrieve number of weekdays between two dates via Aspose.Cells API
// Tags: Aspose.Cells NETWORKDAYS C# implementation | workday count in Excel via Aspose.Cells API | business days calculation in .NET workbook | Excel formula calculation using Aspose.Cells | C# compute weekdays with Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes start and end dates to cells A1 and B1, applies the built‑in NETWORKDAYS formula to cell C1 to determine the number of working days, forces formula evaluation, and saves the file as an XLSX using Aspose.Cells for .NET.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate example dates.
            sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 2)); // Monday
            sheet.Cells["B1"].PutValue(new DateTime(2023, 1, 10)); // Tuesday

            // Use the built‑in NETWORKDAYS function to count workdays between the two dates.
            sheet.Cells["C1"].Formula = "NETWORKDAYS(A1,B1)";

            // Force calculation of formulas.
            workbook.CalculateFormula();

            // Save the workbook.
            workbook.Save("WorkDaysExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
