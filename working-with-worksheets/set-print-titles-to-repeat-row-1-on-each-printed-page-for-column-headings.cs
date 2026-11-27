// Title: Set the first worksheet row as a repeating print title on every printed page with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# repeat header row on each printed page | how to set print title rows in Excel using Aspose.Cells .NET | C# PageSetup.PrintTitleRows example for Excel workbook | Aspose.Cells set print titles for worksheet pages | repeat first row as print title when saving XLSX with Aspose.Cells
// Tags: Aspose.Cells PageSetup PrintTitleRows | C# repeat header row Excel print | Excel worksheet print titles Aspose.Cells | set print title rows .NET | XLSX print title rows Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, adds a header row with sample data, configures the first row as a print title that repeats on every printed page using the PageSetup.PrintTitleRows property, and saves the workbook as an XLSX file.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Quantity");

            for (int i = 2; i <= 50; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue("General");
                sheet.Cells[$"C{i}"].PutValue(10 + i);
                sheet.Cells[$"D{i}"].PutValue(i * 2);
            }

            // Set the first row as print title (repeat on each printed page)
            // The format is "startRow:endRow"
            sheet.PageSetup.PrintTitleRows = "1:1";

            // Optional: you can also set print title columns if needed
            // sheet.PageSetup.PrintTitleColumns = "A:A";

            // Save the workbook
            string outputPath = "PrintTitleRowsDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
