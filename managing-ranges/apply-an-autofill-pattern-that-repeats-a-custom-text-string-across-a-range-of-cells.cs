// Title: How to use Aspose.Cells AutoFill to repeat a custom string across a range of cells in C#
// Common Searches: Aspose.Cells C# autofill same text in a column | repeat a custom value across A1:A20 using Aspose.Cells AutoFill | C# copy cell value to multiple cells with Aspose.Cells AutoFillType.Copy | populate Excel range with identical string using Aspose.Cells in .NET | create source and destination ranges for AutoFill in Aspose.Cells C#
// Tags: AutoFill Copy type Aspose.Cells C# | repeat string across Excel range Aspose.Cells | CreateRange for source and destination Aspose.Cells | fill identical cells Excel .NET | Aspose.Cells repeat custom text

using Aspose.Cells;
using System;

// The example creates a new workbook, writes a custom string to cell A1, defines a single‑cell source range and a destination range (A1:A20), then uses AutoFill with the Copy type to duplicate the text across the entire range before saving the file as AutofillPattern.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the custom text to repeat
            string customText = "CustomString";

            // Place the custom text in the first cell of the target range (e.g., A1)
            worksheet.Cells["A1"].PutValue(customText);

            // Define the source range (single cell containing the custom text)
            Aspose.Cells.Range sourceRange = worksheet.Cells.CreateRange("A1", "A1");

            // Define the destination range where the pattern will be repeated (e.g., A1:A20)
            Aspose.Cells.Range fillRange = worksheet.Cells.CreateRange("A1", "A20");

            // Apply autofill with Copy type to repeat the custom text across the range
            sourceRange.AutoFill(fillRange, AutoFillType.Copy);

            // Save the workbook to a file
            workbook.Save("AutofillPattern.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
