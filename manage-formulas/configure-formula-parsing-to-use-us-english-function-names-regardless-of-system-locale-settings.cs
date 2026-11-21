// Title: Configure Aspose.Cells .NET workbook to always use US English formula function names
// Common Searches: how to force Aspose.Cells to use English function names in C# | set workbook.Settings.CultureInfo to en-US for formula parsing Aspose.Cells | override system locale for Excel formulas when using Aspose.Cells .NET | Aspose.Cells formula language independent of Windows regional settings | C# example to enforce US English formulas in Aspose.Cells workbook
// Tags: Aspose.Cells formula parsing culture en-US | force English formula names in Aspose.Cells | disable locale based formula translation Aspose.Cells | C# set workbook formula language | Excel function name localization Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The sample creates a Workbook, sets workbook.Settings.CultureInfo to "en-US" so that all formulas use US English function names regardless of the system locale, adds data, applies a SUM formula using the English name, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Configure formula parsing to always use US English function names
        workbook.Settings.CultureInfo = new CultureInfo("en-US");

        // Example: add some data and a formula using US English function name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "SUM(A1:A2)";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
