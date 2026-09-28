// Title: Use Aspose.Cells for .NET to autofill a custom string down a column (A1:A15)
// AI Prompts: Generate C# code that writes a given text to cell A1 and uses Aspose.Cells AutoFill to copy it through cells A1 to A15. | Show how to create source and destination Range objects and apply AutoFillType.Copy to repeat a value across a column in an Excel workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# autofill same value from A1 to A15 example | How to copy a single cell value to a range using AutoFillType.Copy in Aspose.Cells .NET | Repeat custom string in an Excel column with Aspose.Cells AutoFill in C#
// Tags: Aspose.Cells AutoFillType.Copy usage | C# repeat value across Excel range with Aspose.Cells | Create source and destination Range Aspose.Cells | populate Excel column with identical text Aspose.Cells | Aspose.Cells autofill custom string to column

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, writes a custom string to cell A1, defines source and destination ranges (A1 and A1:A15), applies AutoFill with AutoFillType.Copy to repeat the string down the column, and saves the file as AutofillCustomText.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Custom text to repeat
            string customText = "MyCustomString";

            // Place the custom text in the starting cell (A1)
            sheet.Cells["A1"].PutValue(customText);

            // Define the source range (the cell with the initial value)
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1", "A1");

            // Define the destination range where the pattern will be repeated (A1:A15)
            AsposeRange destinationRange = sheet.Cells.CreateRange("A1", "A15");

            // Apply autofill to copy the custom text across the destination range
            sourceRange.AutoFill(destinationRange, AutoFillType.Copy);

            // Save the workbook to a file
            string outputPath = "AutofillCustomText.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
