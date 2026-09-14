// Title: Configure Aspose.Cells workbook to use the 1904 (Mac) date system and validate day differences with an Excel formula
// AI Prompts: Enable the 1904 date system by setting workbook.Settings.Date1904 = true, write two DateTime values into cells A1 and A2, assign the formula =A2-A1 to cell A3, force formula calculation, and retrieve the Excel-computed day count. | After switching to the Mac style date system, compare the Excel formula result for the date difference with the .NET TotalDays calculation and output both values.
// Common Searches: Aspose.Cells how to switch workbook to 1904 date system in C# | compare Excel date subtraction result with .NET TotalDays using Aspose.Cells | verify date calculations after enabling Mac style date system in Aspose.Cells | sample code for setting Settings.Date1904 and calculating days between dates | difference between dates in Excel formula vs .NET when using 1904 date system
// Tags: set workbook Settings.Date1904 Aspose.Cells | calculate date difference with Excel formula C# | verify 1904 date system calculation .NET | write DateTime values to cells Aspose.Cells | force formula evaluation Aspose.Cells workbook

using Aspose.Cells;
using System;

// The example creates a new Workbook, activates the 1904 (Mac) date system via Settings.Date1904, writes two DateTime values to cells A1 and A2, places the formula =A2-A1 in A3, forces formula calculation, retrieves the Excel-computed day difference, compares it with the .NET TotalDays result, prints both values, and saves the file as DateSystem1904.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's date system to 1904 (Mac style)
        workbook.Settings.Date1904 = true;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample dates for verification
        DateTime date1 = new DateTime(2020, 1, 1);
        DateTime date2 = new DateTime(2020, 12, 31);

        // Write the sample dates into cells
        sheet.Cells["A1"].PutValue(date1);
        sheet.Cells["A2"].PutValue(date2);

        // Use an Excel formula to calculate the difference in days
        sheet.Cells["A3"].Formula = "=A2-A1";

        // Force calculation of formulas
        workbook.CalculateFormula();

        // Retrieve the result calculated by Excel
        double excelDifference = sheet.Cells["A3"].DoubleValue;

        // Verify the calculation using .NET
        double dotNetDifference = (date2 - date1).TotalDays;

        // Output verification results
        Console.WriteLine($"Date system set to 1904: {workbook.Settings.Date1904}");
        Console.WriteLine($"Date1 (A1): {date1:d}");
        Console.WriteLine($"Date2 (A2): {date2:d}");
        Console.WriteLine($"Excel calculated difference (A3): {excelDifference} days");
        Console.WriteLine($".NET calculated difference: {dotNetDifference} days");

        // Save the workbook (optional)
        workbook.Save("DateSystem1904.xlsx");
    }
}
