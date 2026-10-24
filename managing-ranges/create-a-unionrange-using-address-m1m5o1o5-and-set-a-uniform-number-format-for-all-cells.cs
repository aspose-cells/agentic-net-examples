// Title: Create a union range (M1:M5, O1:O5) and apply a custom number format with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create union range across non‑adjacent columns and set number format | Apply the same custom number format to multiple disjoint cell ranges using Aspose.Cells | C# example for formatting cells M1:M5 and O1:O5 with a two‑decimal format in Aspose.Cells | How to use StyleFlag to apply a number format to a union range in Aspose.Cells .NET
// Tags: union range creation Aspose.Cells .NET | custom number format style Aspose.Cells | StyleFlag NumberFormat property C# | non‑contiguous cell styling Aspose.Cells | save workbook with formatted union range Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, defines a union range covering M1:M5 and O1:O5, builds a style with the custom number format "0.00", applies this style to the union range using a StyleFlag with NumberFormat enabled, and saves the workbook as UnionRangeNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Create a union range that includes M1:M5 and O1:O5
            var unionRange = worksheet.Cells.CreateRange("M1:M5,O1:O5");

            // Define a style with the desired number format (e.g., two decimal places)
            var style = workbook.CreateStyle();
            style.Custom = "0.00";

            // Apply the number format style to all cells in the union range
            var flag = new StyleFlag();
            flag.NumberFormat = true; // Correct property for number formatting
            unionRange.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("UnionRangeNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
