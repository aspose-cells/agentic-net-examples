// Title: C# – Freeze the Top Three Rows in an Excel Worksheet with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells for .NET to freeze the first three rows of a worksheet (no columns frozen) by calling Worksheet.FreezePanes(3, 0, 3, 0) and saving the result as FreezeTopThreeRows.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze top rows Excel | C# freeze first three rows | Aspose.Cells worksheet freeze example | Excel freeze panes .NET
// Common Searches: how to freeze the first three rows using Aspose.Cells | Aspose.Cells C# freeze panes without columns | freeze top rows Excel programmatically .NET | sample code for Worksheet.FreezePanes in C# | Aspose.Cells freeze pane tutorial
// Developer Intent: Apply a freeze‑pane setting that locks the top three rows while allowing all columns to scroll.
// Use Cases: Keep header or summary rows visible while scrolling through large reports. | Generate spreadsheets where the first three rows contain titles, dates, and filters that must stay static. | Apply identical row‑freeze settings across multiple worksheets in an automated workbook creation process.
// AI Prompts: Generate C# code that freezes the first N rows in an Excel file using Aspose.Cells, where N is a variable. | Show how to freeze both rows and columns together with error handling in Aspose.Cells for .NET. | Explain each parameter of Worksheet.FreezePanes and how to remove a freeze‑pane later.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to use Aspose.Cells for .NET to freeze the first three rows of a worksheet (no columns frozen) by calling Worksheet.FreezePanes(3, 0, 3, 0) and saving the result as FreezeTopThreeRows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze the top three rows (no columns are frozen)
                // Parameters: row index, column index, number of frozen rows, number of frozen columns
                worksheet.FreezePanes(3, 0, 3, 0);

                // Save the workbook to a file
                string outputPath = "FreezeTopThreeRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
