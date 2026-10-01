// Title: Refresh every PivotTable in an Excel workbook after bulk data import using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, iterates through all worksheets, refreshes each PivotTable's data source, recalculates its values, and saves the workbook to a new file. | Create a C# routine that validates the input Excel path, opens the workbook, and programmatically calls RefreshData and CalculateData on every PivotTable to reflect recent data changes. | Write error‑handling logic for a C# Aspose.Cells script that processes multiple worksheets, refreshes all PivotTables, and logs any exceptions without terminating the application.
// Common Searches: C# Aspose.Cells how to programmatically refresh all pivot tables after updating source data | Refresh PivotTable collection in each worksheet using Aspose.Cells .NET | Aspose.Cells refresh data and recalculate pivot tables before saving workbook | Bulk data import then update pivot tables with Aspose.Cells C# example | Iterate worksheets and refresh pivot tables in an Excel file using Aspose.Cells
// Tags: Aspose.Cells RefreshData API | CalculateData method for PivotTable | C# bulk import Excel pivot synchronization | Workbook.Save after pivot refresh | Validate input Excel path Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example verifies the input Excel file, loads it with Aspose.Cells, loops through each worksheet and its PivotTable collection, calls RefreshData and CalculateData on every pivot table to align with newly imported data, and then saves the updated workbook while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the data and pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of pivot tables on the current worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Refresh each pivot table to reflect the newly imported data
                foreach (PivotTable pivotTable in pivotTables)
                {
                    // Refresh the data source of the pivot table
                    pivotTable.RefreshData();

                    // Recalculate the pivot table values
                    pivotTable.CalculateData();

                    // Note: Aspose.Cells does not expose a RefreshDataOnOpening property for PivotTable.
                    // The pivot tables will be refreshed when the workbook is opened if the data source is up‑to‑date.
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
