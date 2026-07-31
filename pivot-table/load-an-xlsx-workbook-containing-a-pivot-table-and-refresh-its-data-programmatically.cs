// Title: Programmatically refresh all PivotTables in an XLSX workbook with Aspose.Cells for .NET (C#)
// Description: C# example that loads an existing XLSX file using Aspose.Cells, calls Worksheets.RefreshPivotTables to update every PivotTable across all worksheets, and saves the refreshed workbook. Perfect for automating data updates in reports, dashboards, and scheduled jobs.
// Keywords: Aspose.Cells | RefreshPivotTables | C# Excel pivot refresh | update PivotTable programmatically | load and save XLSX Aspose | Excel pivot cache refresh .NET | automate Excel pivot update | Aspose.Cells API | Excel automation C# | refresh pivot tables .NET
// Common Searches: Aspose.Cells refresh pivot tables C# | How to programmatically refresh PivotTable in XLSX using .NET | Refresh all PivotTables in a workbook with Aspose.Cells | C# code to update Excel pivot cache | Automate pivot table refresh Aspose.Cells | RefreshPivotTables method example
// Developer Intent: Refresh the data of every PivotTable in an existing XLSX workbook and persist the changes by saving the file.
// Use Cases: Refresh pivot tables after source data changes before distributing the report. | Automate nightly refresh of financial PivotTables as part of a scheduled job. | Integrate pivot refresh into a web service that generates up‑to‑date Excel files for end users. | Regenerate monthly sales dashboards with the latest figures. | Update financial models before exporting to PDF or other formats.
// AI Prompts: Generate C# code with Aspose.Cells that refreshes a specific PivotTable by name and saves the workbook. | Provide error‑handling examples for RefreshPivotTables when the source data range is missing. | Show how to refresh pivot tables while preserving existing worksheet formatting and macros using Aspose.Cells. | Write C# code to refresh only selected PivotTables by index with Aspose.Cells. | Explain how to log the refresh operation and handle exceptions for missing source data.

using System;
using Aspose.Cells;

namespace RefreshPivotTableExample
{
    // C# example that loads an existing XLSX file using Aspose.Cells, calls Worksheets.RefreshPivotTables to update every PivotTable across all worksheets, and saves the refreshed workbook. Perfect for automating data updates in reports, dashboards, and scheduled jobs.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file that contains a PivotTable
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Refresh all PivotTables in all worksheets of the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook to a new file (or overwrite the original)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Pivot tables refreshed and workbook saved to: " + outputPath);
        }
    }
}
