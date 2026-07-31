// Title: C# – Update PivotTable Data Source to a Union Range Across Worksheets and Refresh with Aspose.Cells
// Description: Loads an Excel workbook, locates the first PivotTable, changes its source to a union of ranges (e.g., Sheet1!A1:C10 and Sheet2!A1:C10) using PivotTable.ChangeDataSource, then calls RefreshData and CalculateData to apply the new data before saving the file.
// Keywords: Aspose.Cells | C# PivotTable union range | ChangeDataSource method | RefreshData | CalculateData | multiple worksheet source | Excel automation .NET | update PivotTable source | GitHub code example | sample code
// Common Searches: Aspose.Cells set union data source for PivotTable | Refresh PivotTable after changing source in C# | Add another worksheet range to existing PivotTable | ChangeDataSource with multiple ranges example | PivotTable data source union Aspose.Cells
// Developer Intent: Programmatically replace a PivotTable's single‑range source with a union of ranges on different worksheets and refresh the table so the new data is reflected in the report.
// Use Cases: Automated financial reporting that expands a PivotTable to include data from a newly added worksheet. | Dynamic dashboards where the data range grows across multiple sheets and the PivotTable must stay up‑to‑date. | Batch processing of Excel files that need their PivotTables re‑sourced and recalculated without manual intervention.
// AI Prompts: Generate C# code using Aspose.Cells to change a PivotTable's data source to a union of two worksheet ranges and refresh it. | Explain how PivotTable.ChangeDataSource accepts an array of range strings for union sources in Aspose.Cells. | Create robust error handling for missing input files and absent PivotTables when updating a PivotTable's source.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, locates the first PivotTable, changes its source to a union of ranges (e.g., Sheet1!A1:C10 and Sheet2!A1:C10) using PivotTable.ChangeDataSource, then calls RefreshData and CalculateData to apply the new data before saving the file.
class UpdatePivotUnionRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the first worksheet that contains a PivotTable
            Worksheet pivotSheet = null;
            PivotTable pivotTable = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.PivotTables.Count > 0)
                {
                    pivotSheet = ws;
                    pivotTable = ws.PivotTables[0];
                    break;
                }
            }

            if (pivotTable == null)
            {
                Console.WriteLine("No PivotTable found in the workbook.");
                return;
            }

            // Define the new union data source ranges
            string[] newDataSource = new string[]
            {
                "Sheet1!A1:C10",
                "Sheet2!A1:C10"
            };

            // Change the PivotTable's data source to the union of the ranges
            // Aspose.Cells expects an array of range strings for union sources
            pivotTable.ChangeDataSource(newDataSource);

            // Refresh the PivotTable to reflect the new data source
            pivotTable.RefreshData();      // Gather data from the new source
            pivotTable.CalculateData();    // Recalculate the layout

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
