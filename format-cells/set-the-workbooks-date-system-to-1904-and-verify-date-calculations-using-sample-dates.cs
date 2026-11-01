// Title: Set workbook to the 1904 date system and verify date calculations with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook to 1904 date system | calculate days between two dates after enabling 1904 date system in Aspose.Cells | retrieve Excel serial numbers for dates using Aspose.Cells .NET | verify date calculations when workbook uses 1904 date system with Aspose.Cells | save Excel file with 1904 date system using Aspose.Cells C# example
// Tags: workbook date system 1904 Aspose.Cells | date difference formula Aspose.Cells C# | excel serial number extraction Aspose.Cells | save workbook with custom date system Aspose.Cells | formula evaluation after changing date system Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, switches its date system to 1904 via Settings.Date1904, writes two dates to cells A1 and A2, computes the day difference with an Excel formula in A3, forces formula calculation, reads both the calculated difference and the underlying serial numbers, prints the results, and saves the file as DateSystem1904.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the workbook's date system to 1904 (true = 1904 system, false = 1900 system)
        workbook.Settings.Date1904 = true;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample dates to test
        DateTime date1 = new DateTime(2020, 1, 1);
        DateTime date2 = new DateTime(2020, 1, 31);

        // Write dates into cells
        cells["A1"].PutValue(date1);
        cells["A2"].PutValue(date2);

        // Calculate the difference in days using Aspose.Cells formula
        cells["A3"].Formula = "=A2-A1";

        // Force calculation of formulas
        workbook.CalculateFormula();

        // Retrieve the calculated difference
        double dayDifference = cells["A3"].DoubleValue;

        // Output results
        Console.WriteLine("Date System: 1904");
        Console.WriteLine($"Date 1 (A1): {date1:d}");
        Console.WriteLine($"Date 2 (A2): {date2:d}");
        Console.WriteLine($"Difference (A3): {dayDifference} days");

        // Verify internal serial numbers (Excel stores dates as numbers)
        double serial1 = cells["A1"].DoubleValue;
        double serial2 = cells["A2"].DoubleValue;
        Console.WriteLine($"Serial number for A1: {serial1}");
        Console.WriteLine($"Serial number for A2: {serial2}");

        // Save the workbook (optional, demonstrates that the setting persists)
        workbook.Save("DateSystem1904.xlsx");
    }
}
