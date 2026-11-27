// Title: Create an Excel workbook with frozen panes, validate pane settings, and export to HTML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# freeze panes at specific cell and export to HTML | How to check frozen pane coordinates with GetFreezedPanes in Aspose.Cells | Verify frozen rows and columns are preserved in HTML output from Aspose.Cells | C# sample for exporting workbook with split panes to HTML using HtmlSaveOptions | Detect freeze pane markers in generated HTML from Aspose.Cells
// Tags: Aspose.Cells freeze panes HTML export | C# GetFreezedPanes validation | HtmlSaveOptions preserve frozen panes | Excel workbook frozen rows columns verification | Aspose.Cells HTML freeze pane detection

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, populates it with sample data, freezes panes at cell C3 (2 rows and 2 columns), uses GetFreezedPanes to confirm the freeze settings, saves the workbook as XLSX, exports it to HTML with HtmlSaveOptions, reads the generated HTML file, and checks for freeze‑pane markers to ensure the frozen rows and columns are correctly reflected in the HTML output.
class Program
{
    static void Main()
    {
        // Define file names
        string workbookPath = Path.Combine(Directory.GetCurrentDirectory(), "FrozenPaneWorkbook.xlsx");
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "FrozenPaneWorkbook.html");

        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // 2. Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // 3. Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
            sheet.FreezePanes("C3", 2, 2);

            // 4. Verify that the freeze pane settings are as expected using GetFreezedPanes
            bool hasFreeze = sheet.GetFreezedPanes(out int freezeRow, out int freezeColumn, out int frozenRows, out int frozenColumns);
            Console.WriteLine($"GetFreezedPanes returned: {hasFreeze}");
            Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
            Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");

            // Expected values
            int expectedRow = 2;          // zero‑based index for cell C3
            int expectedColumn = 2;
            int expectedFrozenRows = 2;
            int expectedFrozenColumns = 2;

            bool verificationPassed =
                hasFreeze &&
                freezeRow == expectedRow &&
                freezeColumn == expectedColumn &&
                frozenRows == expectedFrozenRows &&
                frozenColumns == expectedFrozenColumns;

            Console.WriteLine($"Freeze pane verification {(verificationPassed ? "PASSED" : "FAILED")}.");

            // 5. Save the workbook as XLSX (optional, just to have a binary version)
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {workbookPath}");

            // 6. Export the workbook to HTML
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Ensure that the HTML export keeps the frozen panes (default behavior)
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"HTML exported to: {htmlPath}");

            // 7. Verify that the generated HTML contains markers for frozen panes
            string htmlContent = File.ReadAllText(htmlPath);
            // Aspose.Cells adds a div with class "asposegrid" and a script that sets scroll positions.
            // A reliable way is to look for the "freeze" keyword in the generated script.
            bool htmlContainsFreeze = htmlContent.IndexOf("freeze", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                      htmlContent.IndexOf("split", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine($"HTML freeze pane marker {(htmlContainsFreeze ? "found" : "not found")}.");

            // Final result
            if (verificationPassed && htmlContainsFreeze)
            {
                Console.WriteLine("Overall verification SUCCESS: frozen panes are correctly reflected in HTML.");
            }
            else
            {
                Console.WriteLine("Overall verification FAILURE: frozen pane settings are not correctly reflected.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
