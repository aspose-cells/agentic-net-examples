// Title: Aspose.Cells C# – Create Custom View "ReportView", Freeze Header Row, Preserve Worksheet View
// Description: Demonstrates how to add a custom view named ReportView, freeze the first row as a header, and retain the original worksheet view (Normal, PageLayout, etc.) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# FreezePanes | custom view Aspose.Cells | preserve worksheet view type | freeze header row .NET | ReportView Excel Aspose
// Common Searches: Aspose.Cells freeze first row and keep view mode | how to create a custom view in Aspose.Cells C# | preserve worksheet view after freezing panes | save Excel file with custom view name using Aspose.Cells
// Developer Intent: Create a workbook, add a custom view called ReportView, freeze the header row, keep the original view setting, and save the file.
// Use Cases: Generate reports where the top row stays visible while the sheet remains in the user's preferred view (Normal, PageLayout, etc.). | Programmatically modify worksheets without altering the existing view configuration for end‑users. | Export data to Excel with a named custom view that includes frozen headers for easier navigation.
// AI Prompts: Write C# code with Aspose.Cells to create a custom view named ReportView, freeze the first row, preserve the original ViewType, and save the workbook. | Show how to use FreezePanes in Aspose.Cells while keeping the worksheet's view mode unchanged before exporting. | Provide an Aspose.Cells .NET example that adds a custom view, freezes header rows, restores the initial view type, and saves as ReportView.xlsx.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomViewExample
{
    // Demonstrates how to add a custom view named ReportView, freeze the first row as a header, and retain the original worksheet view (Normal, PageLayout, etc.) using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Preserve the current view type (e.g., Normal, PageBreakPreview, PageLayout)
                ViewType originalView = worksheet.ViewType;

                // Freeze the first row (header)
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Row index 1 corresponds to the second row (A2), freezing 1 row above it.
                worksheet.FreezePanes(1, 0, 1, 0);

                // Restore the original view type to keep the initial view settings
                worksheet.ViewType = originalView;

                // Define output file path
                string outputPath = "ReportView.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
