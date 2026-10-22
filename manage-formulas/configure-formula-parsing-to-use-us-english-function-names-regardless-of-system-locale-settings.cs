// Title: How to force Aspose.Cells to use US English function names for formula parsing in C#
// Common Searches: Aspose.Cells C# set formula language to English regardless of system locale | force English function names in Excel formulas with Aspose.Cells | configure workbook.Settings.CultureInfo to en-US for formula parsing | override locale for Aspose.Cells formula evaluation in .NET | use US English Excel functions in Aspose.Cells workbook C# example
// Tags: Aspose.Cells formula culture en-us | C# workbook Settings CultureInfo for formulas | English function names in Aspose.Cells | Excel formula localization with Aspose.Cells | override system locale for formula parsing | set workbook culture for formula evaluation

using Aspose.Cells;
using System.Globalization;

// The example creates a Workbook, sets workbook.Settings.CultureInfo to "en-US" so that all formulas use US English function names (e.g., SUM), adds sample values, assigns a formula, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Configure formula parsing to always use US English function names
        workbook.Settings.CultureInfo = new CultureInfo("en-US");

        // Example usage: add some values and a formula that uses US English function name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "SUM(A1:A2)"; // "SUM" is the US English name

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
