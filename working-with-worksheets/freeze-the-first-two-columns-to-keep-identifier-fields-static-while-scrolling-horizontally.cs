// Title: How to freeze the first two columns (A and B) in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze only columns A and B while keeping all rows scrollable. | Show how to determine the used range of a worksheet and apply the FreezePanes method to lock the first two columns in a .NET workbook. | Provide a complete example that creates a workbook, adds header rows, freezes the first two columns, and saves the result as an .xlsx file.
// Common Searches: Aspose.Cells C# freeze columns A and B without freezing rows | How to use FreezePanes to lock the first two columns in an Excel file with .NET | Programmatically set frozen panes for columns only using Aspose.Cells | C# example for freezing specific columns in an Excel worksheet with Aspose | Freeze first two columns in generated Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells FreezePanes column freezing | C# Excel column freeze with Aspose | set frozen panes for columns .NET | Aspose.Cells worksheet FreezePanes usage | generate .xlsx with frozen columns using Aspose

using System;
using Aspose.Cells;

// The sample creates a new workbook, writes header values, computes the worksheet's used range, and calls FreezePanes(0, 2, totalRows, totalColumns) to lock columns A and B while leaving rows unfrozen, then saves the file as FrozenColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Sample header data
            sheet.Cells["A1"].PutValue("Identifier");
            sheet.Cells["B1"].PutValue("Code");
            sheet.Cells["C1"].PutValue("Name");
            sheet.Cells["D1"].PutValue("Description");

            // Freeze the first two columns (A and B)
            // FreezePanes(row, column, totalRows, totalColumns)
            // row = 0 (no rows frozen), column = 2 (freeze first two columns)
            // totalRows/totalColumns define the pane size; using max data range ensures full sheet is considered
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(0, 2, totalRows, totalColumns);

            // Save the workbook
            string outputPath = "FrozenColumns.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
