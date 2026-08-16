// Title: C# Load Excel Workbook with Aspose.Cells and Skip Hidden Rows (AutoFilter)
// Description: Demonstrates how to configure Aspose.Cells LoadOptions with AutoFilter to ignore rows hidden by an existing filter when opening an Excel file, then iterate only visible rows and optionally save the workbook.
// Keywords: Aspose.Cells LoadOptions AutoFilter | C# skip hidden rows Excel | Aspose.Cells ignore filtered rows | .NET read visible rows Excel | Cells.IsRowHidden example | load workbook without hidden rows
// Common Searches: Aspose.Cells load workbook without hidden rows | C# hide filtered rows when reading Excel | How to ignore AutoFilter hidden rows in Aspose.Cells | Read only visible rows from Excel using Aspose.Cells .NET | LoadOptions.AutoFilter usage example
// Developer Intent: Open an Excel file with Aspose.Cells, automatically exclude rows hidden by an AutoFilter, and process only the visible data rows.
// Use Cases: Extract data from a filtered sheet while ignoring hidden rows. | Export or copy only visible rows to another workbook. | Run calculations or analytics on rows that meet the filter criteria.
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, enables AutoFilter, and iterates only over rows that are not hidden. | Show how to use LoadOptions.AutoFilter and Cells.IsRowHidden to skip filtered‑out rows in a .NET application. | Explain the steps to configure Aspose.Cells so hidden rows are excluded during workbook loading.

using System;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells LoadOptions with AutoFilter to ignore rows hidden by an existing filter when opening an Excel file, then iterate only visible rows and optionally save the workbook.
class Program
{
    static void Main()
    {
        // Create load options and enable AutoFilter.
        // This tells Aspose.Cells to apply any existing autofilter
        // in the source file and hide rows that do not meet the filter criteria.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFilter = true;

        // Load the workbook with the specified options.
        // Hidden rows (as determined by the autofilter) will be ignored.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Iterate through all rows that contain data.
        // Process only rows that are not hidden.
        int maxRow = cells.MaxDataRow;
        for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
        {
            if (!cells.IsRowHidden(rowIndex))
            {
                // Example processing: output the value of the first column.
                Console.WriteLine($"Row {rowIndex + 1}: {cells[rowIndex, 0].StringValue}");
            }
        }

        // Save the workbook if further actions are required.
        workbook.Save("output.xlsx");
    }
}
