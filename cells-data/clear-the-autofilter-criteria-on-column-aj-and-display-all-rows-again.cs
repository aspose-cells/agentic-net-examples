using System;
using Aspose.Cells;

class ClearAutoFilterOnColumnAJ
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the last row that contains data
        int lastDataRow = worksheet.Cells.MaxDataRow; // zero‑based index

        // Ensure the AutoFilter range includes column AJ (column index 35)
        // Build the range string like "A1:AJ{lastRow+1}"
        string autoFilterRange = $"A1:AJ{lastDataRow + 1}";
        worksheet.AutoFilter.Range = autoFilterRange;

        // Clear any filter criteria on column AJ and display all rows again
        worksheet.AutoFilter.ShowAll();

        // Refresh the filter to apply the change immediately
        worksheet.AutoFilter.Refresh();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}