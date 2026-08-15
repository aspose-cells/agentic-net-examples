// Title: C# – Refresh All Pivot Tables in an Aspose.Cells Workbook After Bulk Data Import
// Description: Loads an existing workbook, simulates bulk data import by updating a range of cells, calls Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and saves the result. Demonstrates error handling and performance‑friendly random data generation.
// Keywords: Aspose.Cells | C# | RefreshPivotTables | pivot table refresh | bulk data import | Workbook.Worksheets.RefreshPivotTables | programmatic pivot update | Excel automation | data‑driven reporting
// Common Searches: how to refresh all pivot tables with Aspose.Cells C# | RefreshPivotTables method example | update pivot tables after bulk import Aspose | Aspose.Cells refresh pivot tables programmatically | C# code to recalculate pivot tables in Excel file
// Developer Intent: Recalculate every pivot table in a workbook so it reflects newly imported data before saving.
// Use Cases: Automated reporting pipelines that modify source data and need up‑to‑date pivot summaries. | Template‑based workbooks where large data sets are injected and pivot‑driven dashboards must stay accurate. | Scheduled jobs that ingest sales figures, update cells, and refresh all pivot tables to generate fresh reports.
// AI Prompts: Write C# code that updates a range of cells in an Aspose.Cells workbook and then refreshes all pivot tables. | Show how to handle exceptions when calling RefreshPivotTables after a bulk data import. | Explain how to refresh pivot tables in selected worksheets only using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPivotRefreshDemo
{
    // Loads an existing workbook, simulates bulk data import by updating a range of cells, calls Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and saves the result. Demonstrates error handling and performance‑friendly random data generation.
    public class RefreshAllPivotTables
    {
        public static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook that contains pivot tables
                Workbook workbook = new Workbook(inputPath);

                // Assume bulk data import modifies a range of cells in the first worksheet
                Worksheet dataSheet = workbook.Worksheets[0];
                Cells cells = dataSheet.Cells;

                // Use a single Random instance for better performance
                Random rnd = new Random();

                // Example bulk import: update sales values for rows 2 to 101 (zero‑based index)
                for (int row = 1; row <= 100; row++) // row 1 = Excel row 2
                {
                    // Update column B (index 1) with a new random sales figure
                    cells[row, 1].PutValue(rnd.Next(1000, 5000));
                }

                // After data changes, refresh all pivot tables in the workbook
                workbook.Worksheets.RefreshPivotTables();

                // Save the updated workbook
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
