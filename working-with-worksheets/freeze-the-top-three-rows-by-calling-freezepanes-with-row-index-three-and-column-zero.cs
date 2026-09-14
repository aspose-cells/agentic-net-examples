// Title: How to freeze the top three rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to freeze rows 1‑3 of a worksheet while keeping columns unfrozen, then save the file. | Demonstrate calling Worksheet.FreezePanes to lock the first three rows without affecting columns in a .NET application. | Provide a minimal Aspose.Cells example that creates a workbook, freezes the top three rows, and writes the result to FrozenRows.xlsx.
// Common Searches: Aspose.Cells C# freeze first three rows without freezing columns | example of FreezePanes method to lock top rows in .NET Excel workbook | how to programmatically freeze rows 1 to 3 using Aspose.Cells for .NET
// Tags: Aspose.Cells FreezePanes top rows | C# freeze first three rows Excel | worksheet freeze rows without columns | save workbook after applying freeze panes | Excel freeze panes programmatic Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, accesses the first worksheet, freezes the top three rows with Worksheet.FreezePanes(3, 0, 3, 0) while leaving columns unfrozen, and saves the file as FrozenRows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the top three rows (row index 3) and no columns
                // Parameters: row index where split occurs, column index where split occurs,
                // number of rows to freeze, number of columns to freeze
                sheet.FreezePanes(3, 0, 3, 0);

                // Save the workbook to a file
                string outputPath = "FrozenRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
