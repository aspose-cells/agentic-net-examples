// Title: Detect Header‑Only Worksheets with Aspose.Cells for .NET (MaxDataRow = -1, MaxDataColumn > 0)
// Description: C# example that loads an Excel file, scans each worksheet, and uses Cells.MaxDataRow (‑1) together with Cells.MaxDataColumn (>0) to identify sheets that contain only column headers. The worksheet names are printed and the workbook is saved unchanged.
// Keywords: Aspose.Cells header only worksheet | MaxDataRow -1 | MaxDataColumn detection | C# Excel sheet header check | .NET identify empty data rows | Aspose.Cells worksheet analysis
// Common Searches: Aspose.Cells find sheets with only headers | MaxDataRow -1 meaning in Aspose.Cells | Check if Excel worksheet has data rows C# | Detect header‑only worksheets using Aspose.Cells
// Developer Intent: Locate worksheets that consist solely of column headers and list their names.
// Use Cases: Skip header‑only sheets during bulk data processing. | Validate imported workbooks to ensure each sheet contains data beyond the header row. | Flag or log worksheets lacking data rows for data‑quality audits.
// AI Prompts: Generate C# code with Aspose.Cells that writes the names of header‑only worksheets to a text file. | Explain the difference between MaxDataRow = -1 and MaxDataColumn = 0 in Aspose.Cells and how they help distinguish empty sheets from header‑only sheets. | Modify the sample to also detect completely empty worksheets and handle them in a separate branch.

using System;
using Aspose.Cells;

// C# example that loads an Excel file, scans each worksheet, and uses Cells.MaxDataRow (‑1) together with Cells.MaxDataColumn (>0) to identify sheets that contain only column headers. The worksheet names are printed and the workbook is saved unchanged.
class DetectHeaderOnlySheets
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow == -1 means there are no data rows (only possible headers)
            // MaxDataColumn > 0 means at least one column contains a value (the header)
            if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn > 0)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only column headers.");
            }
        }

        // Save the workbook (no modifications made, just demonstrating the save rule)
        workbook.Save("output.xlsx");
    }
}
