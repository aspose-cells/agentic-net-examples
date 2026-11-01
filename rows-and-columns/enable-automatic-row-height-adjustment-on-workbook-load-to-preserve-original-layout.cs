// Title: Automatically auto‑fit row heights for every worksheet when loading an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, loop through each worksheet, and invoke Worksheet.AutoFitRows to automatically adjust row heights based on cell content. | Use Cells.MaxDataRow to find the last populated row in a sheet, then call AutoFitRows from the first row to that row to keep the original layout intact. | After auto‑fitting rows on all worksheets, save the workbook to a new file while preserving formatting and layout.
// Common Searches: how to auto‑fit row heights for all sheets when opening an Excel file with Aspose.Cells C# | Aspose.Cells auto adjust row height after loading workbook .NET | preserve original Excel layout by auto‑sizing rows programmatically in C#
// Tags: Worksheet.AutoFitRows Aspose.Cells .NET | auto‑fit rows after workbook load | adjust row heights based on cell content C# | preserve Excel layout Aspose.Cells | iterate all worksheets auto‑fit rows

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet, determines the last row containing data, applies Worksheet.AutoFitRows to automatically adjust row heights from the first row to the last populated row, and saves the workbook with the updated layout.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last row that contains data
            int lastDataRow = sheet.Cells.MaxDataRow;

            // If there is at least one row with data, auto‑fit the row heights
            if (lastDataRow >= 0)
            {
                // AutoFitRows(startRow, totalRows)
                // startRow = 0 (first row), totalRows = lastDataRow + 1 (to include the last row)
                sheet.AutoFitRows(0, lastDataRow + 1);
            }
        }

        // Save the workbook after adjusting row heights
        workbook.Save("output.xlsx");
    }
}
