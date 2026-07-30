// Title: Refresh Data Connections & Freeze Header Row in Every Worksheet – Aspose.Cells C# Example
// Description: This C# sample loads a workbook, updates all linked data sources, pivot tables and charts using Workbook.RefreshAll, then locks the first row of each sheet with FreezePanes, and saves the result. Includes file‑existence verification and exception handling.
// Keywords: Aspose.Cells | Workbook.RefreshAll | Refresh data connections | FreezePanes | C# .NET | freeze header row | update pivot tables | refresh charts | automate Excel processing | Excel workbook manipulation
// Common Searches: Aspose.Cells refresh all connections C# | C# freeze first row each worksheet Aspose.Cells | Workbook.RefreshAll usage example | How to apply FreezePanes to all sheets in Aspose.Cells | Update external data and lock headers with Aspose.Cells
// Developer Intent: Update every external data link, pivot table and chart, then keep the top row visible across all worksheets.
// Use Cases: Nightly data pull for sales dashboards where refreshed figures must stay aligned with frozen column titles. | Regulatory financial statements that require up‑to‑date data while preserving header visibility for auditors in the US and EU. | Large inventory reports distributed to remote teams, ensuring headers stay in view after automatic data refresh. | Generating printable PDFs where refreshed content and frozen rows guarantee consistent page layout.
// AI Prompts: Write a C# program using Aspose.Cells that refreshes all data connections, then freezes the first two rows of each worksheet. | Provide a robust Aspose.Cells snippet that calls Workbook.RefreshAll, applies FreezePanes with custom row/column offsets, and logs any refresh errors. | Explain the interaction between Workbook.RefreshAll and external data sources, and how FreezePanes can be combined to improve user experience in Excel viewers.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# sample loads a workbook, updates all linked data sources, pivot tables and charts using Workbook.RefreshAll, then locks the first row of each sheet with FreezePanes, and saves the result. Includes file‑existence verification and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "InputWorkbook.xlsx";
            const string outputFile = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Refresh all pivot tables, charts, and data connections (use Workbook.RefreshAll)
                workbook.RefreshAll();

                // Freeze the first row in each worksheet to keep headers visible
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Freeze first row (row index 1 is the first scrollable row, freeze 1 row, no columns)
                    sheet.FreezePanes(1, 0, 1, 0);
                }

                // Save the updated workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook successfully saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
