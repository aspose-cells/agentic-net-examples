// Title: Convert string representations to numbers and dates in an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that populates a worksheet with string values, calls Cells.ConvertStringToNumericValue, and prints the resulting numeric and date values. | Show how to use Aspose.Cells to automatically detect and transform string‑encoded numbers and dates into proper numeric and DateTime cell types. | Demonstrate the steps to save a workbook after converting string cells to their native numeric or date formats with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# convert string cells to numeric values example | How to change string dates to DateTime in an Excel file using Aspose.Cells | C# method to automatically convert numeric strings in a worksheet to numbers with Aspose.Cells | Convert mixed string data to proper types in Aspose.Cells workbook C#
// Tags: Cells.ConvertStringToNumericValue C# | string to numeric conversion Aspose.Cells | date string to DateTime Aspose.Cells | populate worksheet with string values Aspose.Cells | save workbook after type conversion Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, inserts string representations of a number, a date, and a non‑numeric value into cells A1‑C1, invokes Cells.ConvertStringToNumericValue to transform convertible strings into numeric or date types, prints the converted values, and saves the workbook as ConvertedValues.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Populate cells with textual representations
        // These values are initially stored as strings
        cells[0, 0].PutValue("123.45");          // numeric string
        cells[0, 1].PutValue("2023-07-15");      // date string
        cells[0, 2].PutValue("NotANumber");      // non‑numeric string

        // Convert all possible string values to their numeric/date equivalents
        // The method scans the worksheet and changes the cell type where conversion is feasible
        cells.ConvertStringToNumericValue();

        // Output the results to verify conversion
        Console.WriteLine("A1 (numeric) : " + cells[0, 0].DoubleValue);          // 123.45
        Console.WriteLine("B1 (date)    : " + cells[0, 1].DateTimeValue);        // 7/15/2023
        Console.WriteLine("C1 (string)  : " + cells[0, 2].StringValue);          // NotANumber

        // Save the workbook (lifecycle save)
        workbook.Save("ConvertedValues.xlsx");
    }
}
