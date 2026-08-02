// Title: C# – Set Row Heights and Freeze the First Two Rows with Aspose.Cells
// Description: Creates a workbook, sets row 1 to 30 pt and row 2 to 45 pt, adds sample text, freezes the top two rows so the custom heights stay visible while scrolling, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | SetRowHeight | FreezePanes | custom row height | freeze rows | Excel worksheet | row height points | freeze panes .NET | Aspose.Cells example
// Common Searches: Aspose.Cells set row height C# | Freeze first rows Aspose.Cells .NET | How to keep custom row heights when freezing panes | C# code to set row height and freeze panes in Excel | Aspose.Cells FreezePanes example with custom heights
// Developer Intent: Generate an Excel workbook, apply specific heights to the first two rows, lock those rows in place, and persist the result.
// Use Cases: Design a report header with larger rows that remain visible during scrolling. | Create a printable template where top rows have custom heights and are frozen. | Build a data‑driven spreadsheet where title rows need distinct height and fixed positioning.
// AI Prompts: Provide a C# Aspose.Cells snippet that sets row 0 to 30 pt, row 1 to 45 pt, and freezes the first two rows. | Explain the interaction between SetRowHeight and FreezePanes in Aspose.Cells for .NET. | Generate code to add sample values, apply custom row heights, freeze the top rows, and save as XLSX using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, sets row 1 to 30 pt and row 2 to 45 pt, adds sample text, freezes the top two rows so the custom heights stay visible while scrolling, and saves the file as XLSX.
    public class SetRowHeightsAndFreezeDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set custom heights for the first two rows (in points)
            // Row 0 (first row) height = 30 points
            cells.SetRowHeight(0, 30);
            // Row 1 (second row) height = 45 points
            cells.SetRowHeight(1, 45);

            // Optionally add some data to visualize the rows
            cells["A1"].PutValue("First row with custom height");
            cells["A2"].PutValue("Second row with custom height");

            // Freeze the first two rows so their custom heights stay visible while scrolling
            // Freeze at row index 2 (the row after the frozen area), column index 0 (no columns frozen)
            worksheet.FreezePanes(2, 0, 2, 0);

            // Determine output path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomRowHeightAndFreeze.xlsx");

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
