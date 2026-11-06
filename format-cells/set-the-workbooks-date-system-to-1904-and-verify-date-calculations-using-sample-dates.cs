// Title: Switch workbook to 1904 date system and compute day difference using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook to 1904 date system and get date difference | calculate days between two dates in Excel file with 1904 date system using Aspose.Cells | C# example for Settings.Date1904 property and formula evaluation in Aspose.Cells | verify date calculations after enabling 1904 date system in Aspose.Cells | Aspose.Cells workbook.Save with 1904 date system example
// Tags: Aspose.Cells Settings.Date1904 usage | C# workbook date system 1904 | Excel date difference formula Aspose.Cells | calculate day interval with Aspose.Cells | save workbook with 1904 date system

using System;
using Aspose.Cells;

// Shows how to enable the 1904 date system in an Aspose.Cells workbook, insert sample dates, apply a formula to compute the day difference, recalculate formulas, and save the file using C#.
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

        // Sample dates
        DateTime date1 = new DateTime(2020, 1, 1);
        DateTime date2 = new DateTime(2020, 1, 15);

        // Place the dates into cells A1 and A2
        cells["A1"].PutValue(date1);
        cells["A2"].PutValue(date2);

        // Add a formula in B1 to calculate the difference in days between the two dates
        cells["B1"].Formula = "A2-A1";

        // Recalculate formulas
        workbook.CalculateFormula();

        // Retrieve the calculated difference
        double dayDifference = cells["B1"].DoubleValue;

        // Output the results
        Console.WriteLine($"Date System: {(workbook.Settings.Date1904 ? "1904" : "1900")}");
        Console.WriteLine($"Date 1 (A1): {date1:d}");
        Console.WriteLine($"Date 2 (A2): {date2:d}");
        Console.WriteLine($"Difference in days (B1): {dayDifference}");

        // Save the workbook (optional)
        workbook.Save("DateSystem1904.xlsx");
    }
}
