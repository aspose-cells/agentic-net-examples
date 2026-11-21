// Title: Set PivotTable.IsExcel2003Compatible to true and refresh the pivot using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set PivotTable.IsExcel2003Compatible before RefreshData | How to enforce 255‑character truncation on pivot fields with Aspose.Cells | Refresh a pivot table after enabling Excel 2003 compatibility in .NET | C# example to truncate pivot values to 255 characters using Aspose.Cells | Load workbook, modify pivot compatibility, and save with Aspose.Cells
// Tags: Aspose.Cells PivotTable IsExcel2003Compatible property | C# enable Excel 2003 compatibility for pivot tables | RefreshData CalculateData Aspose.Cells pivot operation | truncate pivot field values 255 characters Aspose.Cells | load workbook modify pivot save Aspose.Cells workflow

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The code loads an existing workbook, accesses a named PivotTable, sets its IsExcel2003Compatible property to true to enforce the 255‑character limit, refreshes and recalculates the pivot data, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";
            const string pivotName = "MyPivotTable";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the PivotTable by name
                PivotTable pivot = sheet.PivotTables[pivotName];
                if (pivot == null)
                {
                    Console.WriteLine($"Error: PivotTable named '{pivotName}' was not found on the worksheet.");
                    return;
                }

                // Enforce Excel 2003 compatibility (255‑character truncation) before refreshing
                pivot.IsExcel2003Compatible = true;

                // Refresh the PivotTable data and recalculate
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
