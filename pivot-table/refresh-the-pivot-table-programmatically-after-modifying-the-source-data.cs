// Title: Refresh all pivot tables after modifying source data in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh pivot tables after changing source data | How to recalculate pivot cache programmatically using Aspose.Cells .NET | Update a cell value and refresh all pivot tables in an Excel file with Aspose.Cells | Iterate through worksheets to refresh pivot tables using Aspose.Cells API | RefreshData and CalculateData methods example in Aspose.Cells C#
// Tags: Aspose.Cells RefreshData method C# | Aspose.Cells CalculateData API | refresh pivot tables Excel Aspose.Cells | update source data and refresh pivot cache Aspose.Cells | iterate worksheets pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable class resides in this namespace

// The example loads an existing workbook, modifies cell B2 on the "Data" sheet, then loops through every worksheet and each pivot table, calling RefreshData() and CalculateData() to update the pivot cache and recalculate values before saving the workbook.
class PivotTableRefreshExample
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Modify the source data (example: change cell B2 in the sheet named "Data")
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Error: Worksheet 'Data' does not exist.");
                return;
            }
            dataSheet.Cells["B2"].PutValue(12345);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the pivot table's cache with the updated source data
                    pivotTable.RefreshData();

                    // Recalculate the pivot table values
                    pivotTable.CalculateData();
                }
            }

            // Save the workbook with the refreshed pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
