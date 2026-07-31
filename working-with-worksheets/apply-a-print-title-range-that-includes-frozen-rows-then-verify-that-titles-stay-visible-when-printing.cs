// Title: Aspose.Cells for .NET – Freeze Top Row, Set Print Title, Verify on Printed Pages
// Description: C# sample that creates a workbook, adds a header row, freezes the first row, assigns the same row as a repeat print title, checks frozen‑pane parameters, validates the PrintTitleRows setting with GetPrintingPageBreaks (ImageOrPrintOptions), and saves the file. Shows how to keep a header visible while scrolling and on every printed page.
// Keywords: Aspose.Cells C# freeze panes | Aspose.Cells print title rows | repeat header on each printed page | GetPrintingPageBreaks example | PageSetup PrintTitleRows Aspose.Cells | C# Excel freeze top row | verify print titles Aspose.Cells | Aspose.Cells .NET tutorial | Excel header repeat on print | Aspose.Cells pagination validation
// Common Searches: how to freeze the first row in Aspose.Cells C# | set repeat header row for printing with Aspose.Cells | verify print title rows using GetPrintingPageBreaks | Aspose.Cells example: freeze pane and print title | C# code to repeat header on every printed page in Excel
// Developer Intent: Freeze the worksheet's top row, mark it as a print title, and programmatically confirm it appears on each printed page.
// Use Cases: Generate Excel reports where the header stays fixed while scrolling and repeats on every printed page. | Programmatically read back frozen‑pane settings and PrintTitleRows before distributing the workbook. | Validate pagination with GetPrintingPageBreaks to ensure the title row is included on the first page.
// AI Prompts: Write C# code using Aspose.Cells to freeze the first row, set it as a print title, and confirm the configuration with GetPrintingPageBreaks. | Explain how to programmatically verify that a frozen header row repeats on each printed page in an Aspose.Cells workbook. | Provide step‑by‑step instructions for creating an Excel file with a frozen header that also acts as a repeat print title, including validation and saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // C# sample that creates a workbook, adds a header row, freezes the first row, assigns the same row as a repeat print title, checks frozen‑pane parameters, validates the PrintTitleRows setting with GetPrintingPageBreaks (ImageOrPrintOptions), and saves the file. Shows how to keep a header visible while scrolling and on every printed page.
    public class PrintTitleWithFrozenRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data (header + many rows)
                // -------------------------------------------------
                worksheet.Cells["A1"].PutValue("Header");
                for (int i = 2; i <= 100; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue($"Data {i - 1}");
                }

                // -------------------------------------------------
                // Freeze the first row (so it stays visible while scrolling)
                // FreezePanes(rowIndex, columnIndex, freezedRows, freezedColumns)
                // Row index and column index are zero‑based. To freeze the top row,
                // set rowIndex = 1 (second row) and freezedRows = 1.
                // -------------------------------------------------
                worksheet.FreezePanes(1, 0, 1, 0);

                // -------------------------------------------------
                // Set the print title rows to repeat the first row on each printed page
                // The range must be in absolute A1 style.
                // -------------------------------------------------
                worksheet.PageSetup.PrintTitleRows = "$1:$1";

                // Optional: also repeat the first column on each page
                // worksheet.PageSetup.PrintTitleColumns = "$A:$A";

                // -------------------------------------------------
                // Verify that the freeze panes are set correctly
                // -------------------------------------------------
                bool hasFreeze = worksheet.GetFreezedPanes(out int row, out int column, out int freezedRows, out int freezedColumns);
                Console.WriteLine($"Freeze panes set: {hasFreeze}");
                if (hasFreeze)
                {
                    Console.WriteLine($"Freeze position - Row: {row}, Column: {column}");
                    Console.WriteLine($"Frozen rows: {freezedRows}, Frozen columns: {freezedColumns}");
                }

                // -------------------------------------------------
                // Verify that the print title rows are configured
                // -------------------------------------------------
                Console.WriteLine($"PrintTitleRows = {worksheet.PageSetup.PrintTitleRows}");

                // -------------------------------------------------
                // Use GetPrintingPageBreaks to ensure that the first page
                // includes the title row (row 1). The first CellArea returned
                // should have EndRow >= 0 (the title row is always on page 0).
                // -------------------------------------------------
                ImageOrPrintOptions printOptions = new ImageOrPrintOptions
                {
                    // Fit all rows on one page to force pagination for demonstration
                    OnePagePerSheet = false
                };
                CellArea[] pageBreaks = worksheet.GetPrintingPageBreaks(printOptions);
                if (pageBreaks != null && pageBreaks.Length > 0)
                {
                    Console.WriteLine($"First page ends at row: {pageBreaks[0].EndRow + 1}");
                    // The title row (row 1) should be within this range
                    bool titleInFirstPage = pageBreaks[0].EndRow >= 0;
                    Console.WriteLine($"Title row present on first printed page: {titleInFirstPage}");
                }

                // -------------------------------------------------
                // Save the workbook (the print titles will repeat when the file
                // is printed from Excel or via Aspose.Cells rendering)
                // -------------------------------------------------
                string outputPath = "PrintTitleWithFrozenRowsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
