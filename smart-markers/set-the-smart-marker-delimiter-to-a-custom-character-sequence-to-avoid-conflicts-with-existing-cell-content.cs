// Title: How to configure a custom smart marker delimiter in Aspose.Cells for .NET to prevent cell content conflicts
// Common Searches: asp.net change smart marker delimiters in Aspose.Cells | c# set alternative start and end symbols for smart markers | avoid delimiter clash with existing Excel cell text using Aspose.Cells | configure custom smart marker syntax for Excel export in .NET | how to define non‑default smart marker delimiters in C#
// Tags: custom smart marker delimiters Aspose.Cells | smart marker delimiter configuration C# | excel smart marker delimiter conflict resolution | Aspose.Cells set delimiter start end | smart marker syntax customization .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, inserts a smart marker ${CustomerName} into cell A1 using the default delimiters, saves the workbook as CustomSmartMarkerDelimiter.xlsx, and logs success or error messages to the console.
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

            // Insert a smart marker using the default delimiters (${...})
            sheet.Cells["A1"].PutValue("${CustomerName}");

            // Define output file path
            string outputPath = "CustomSmartMarkerDelimiter.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
