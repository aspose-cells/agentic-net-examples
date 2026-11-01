// Title: Refresh all pivot tables and hide every column field item in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file, refreshes each pivot table, marks all items in the column fields as hidden, and saves the result to a new workbook with Aspose.Cells. | Generate a C# program that iterates through all worksheets, calls RefreshData on every pivot table, collapses (hides) all column‑area items, and writes the updated file using Aspose.Cells.
// Common Searches: aspocells c# collapse all column items in pivot table after refresh | programmatically hide pivot column field items in .xlsx using Aspose.Cells | refresh pivot tables and set column items hidden in C# Aspose.Cells example | how to collapse column area of pivot tables in an Excel workbook with Aspose.Cells for .NET | C# code to hide all column field items in every pivot table of a workbook
// Tags: Aspose.Cells refresh pivot data C# | collapse pivot column items Aspose.Cells | hide pivot column field items .xlsx | iterate worksheets pivot tables Aspose.Cells | save modified workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads Report.xlsx, refreshes each pivot table, hides all items in the column fields, and saves the modified workbook as Report_Collapsed.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputFile = "Report.xlsx";
        const string outputFile = "Report_Collapsed.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file '{inputFile}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Refresh the pivot table data
                    pivot.RefreshData();

                    // Collapse (hide) all items in the column area
                    foreach (PivotField colField in pivot.ColumnFields)
                    {
                        foreach (PivotItem item in colField.PivotItems)
                        {
                            // Hide the item to achieve a collapsed view
                            item.IsHidden = true;
                        }
                    }
                }
            }

            // Save the workbook with collapsed column items
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
