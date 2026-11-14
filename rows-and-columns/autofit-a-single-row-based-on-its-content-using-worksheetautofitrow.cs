// Title: Auto‑fit the height of a single worksheet row based on its content using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# auto fit row height for long text in a cell | How to use Worksheet.AutoFitRow to adjust a specific row in .NET | C# code example for auto‑fitting a single row in an Excel workbook with Aspose | Set row height automatically based on cell content using Aspose.Cells | AutoFitRow overload that takes only row index in Aspose.Cells for .NET
// Tags: Worksheet.AutoFitRow row height API | single row auto‑sizing with Aspose.Cells | C# Excel row height auto‑adjustment | auto‑fit row using cell content .NET | AutoFitRow overload row index

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes a long text string into cell A1 of the first worksheet, calls Worksheet.AutoFitRow(0) to automatically adjust the height of that row according to the cell's content, and saves the file as AutoFitRowExample.xlsx.
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

            // Add a long text to cell A1
            sheet.Cells["A1"].PutValue(
                "This is a very long piece of text that will cause the row height to increase when auto‑fitted.");

            // Auto‑fit the first row based on its content
            // Use the overload that accepts only the row index (available in all versions)
            sheet.AutoFitRow(0);

            // Save the workbook
            string outputPath = "AutoFitRowExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
