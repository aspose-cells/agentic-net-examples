// Title: Refresh All Pivot Tables After Updating a Worksheet Table with Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook, change values inside a worksheet ListObject, call Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table that uses the altered data, and save the updated file.
// Keywords: Aspose.Cells | RefreshPivotTables | C# | pivot table refresh | worksheet table update | ListObject | programmatic Excel pivot refresh | Excel automation .NET
// Common Searches: Aspose.Cells refresh pivot tables after data change | C# update Excel table and refresh pivots | How to programmatically refresh all pivots in Aspose.Cells | RefreshPivotTables method example .NET | Refresh all pivot tables in workbook using Aspose.Cells
// Developer Intent: Programmatically refresh every pivot table that depends on a modified worksheet table.
// Use Cases: After bulk updating rows in a data table, automatically refresh linked pivot reports before saving the workbook. | Process a batch of workbooks to change source data and ensure all pivot analyses reflect the new values. | Integrate table modifications and pivot refresh into an ETL pipeline so downstream analytics use the latest figures.
// AI Prompts: Generate C# code that updates a specific ListObject range and then refreshes all pivot tables that reference it using Aspose.Cells. | Show how to refresh pivot tables on a single worksheet instead of the entire workbook with Aspose.Cells for .NET. | Provide a script that iterates through all Excel files in a folder, updates a data table, calls RefreshPivotTables, and saves each file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPivotRefreshDemo
{
    // Shows how to load an Excel workbook, change values inside a worksheet ListObject, call Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table that uses the altered data, and save the updated file.
    public class RefreshPivotTablesAfterTableUpdate
    {
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
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the data table and pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Assume the data table is on the first worksheet (index 0)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Example: update some underlying data in the table (cells A2:B5)
            // In a real scenario, you would locate the ListObject (table) and modify its range.
            dataSheet.Cells["A2"].PutValue("UpdatedItem1");
            dataSheet.Cells["B2"].PutValue(1234);
            dataSheet.Cells["A3"].PutValue("UpdatedItem2");
            dataSheet.Cells["B3"].PutValue(5678);

            // Refresh all pivot tables in the workbook.
            // This will refresh any pivot table that uses the updated data range.
            workbook.Worksheets.RefreshPivotTables();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }
}
