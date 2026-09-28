// Title: Refresh every PivotTable in an Excel workbook sequentially with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, loops through all worksheets, calls RefreshPivotTables() on each sheet, and saves the updated workbook. | Adapt the sample to refresh pivot tables only on worksheets whose names contain "Report" and write each refreshed sheet name to the console. | Wrap the pivot‑refresh loop in try‑catch blocks, log any exceptions, and ensure the workbook is saved even if errors occur.
// Common Searches: Aspose.Cells C# how to refresh all pivot tables in an Excel file | C# loop through worksheets and refresh pivot tables using Aspose.Cells | Refresh pivot tables sequentially before saving workbook with Aspose.Cells .NET | Update pivot cache for multiple sheets Aspose.Cells example
// Tags: Aspose.Cells RefreshPivotTables method | C# iterate worksheets refresh pivot tables | Excel workbook pivot cache synchronization Aspose.Cells | save workbook after pivot refresh Aspose.Cells | sequential pivot table update .NET

using System;
using Aspose.Cells;

namespace RefreshAllPivotTablesDemo
{
    // Loads an Excel workbook, iterates over each worksheet invoking RefreshPivotTables() to synchronize all pivot tables, and saves the refreshed workbook to a new file.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the pivot tables.
            // Replace "input.xlsx" with the path to your source file.
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet and refresh its pivot tables sequentially.
            // This uses the Worksheet.RefreshPivotTables() method.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.RefreshPivotTables();
            }

            // Save the workbook after all pivot tables have been refreshed.
            // The file "output.xlsx" will contain the synchronized data.
            workbook.Save("output.xlsx");
        }
    }
}
