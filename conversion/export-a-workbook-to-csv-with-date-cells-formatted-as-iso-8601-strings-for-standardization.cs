// Title: Export an Excel workbook to CSV with ISO‑8601 date formatting using Aspose.Cells for .NET
// Common Searches: c# aspocells export workbook to csv with iso 8601 date format | how to preserve date formatting as ISO 8601 when saving Excel as CSV using Aspose.Cells | apply custom date style to DateTime cells before CSV export in Aspose.Cells .NET | convert xlsx to csv while converting dates to yyyy-mm-ddThh:mm:ss with Aspose.Cells | Aspose.Cells format datetime cells to ISO 8601 then save as CSV in C#
// Tags: export workbook to CSV with ISO 8601 dates Aspose.Cells | apply custom date style to DateTime cells Aspose.Cells .NET | format Excel cells to yyyy-MM-ddTHH:mm:ss using Aspose.Cells | save XLSX as CSV preserving date format Aspose.Cells | custom style for DateTime cells before CSV export .NET

using Aspose.Cells;
using System;

// The example loads an XLSX workbook, scans every worksheet for DateTime cells, applies a custom style with the pattern "yyyy-MM-ddTHH:mm:ss" to convert dates to ISO‑8601 strings, and then saves the workbook as a CSV file where the formatted dates are written as text.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and format date cells to ISO 8601
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Identify cells that contain DateTime values
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        // Apply ISO 8601 format (e.g., 2023-07-21T15:30:00)
                        Style style = cell.GetStyle();
                        style.Custom = "yyyy-MM-ddTHH:mm:ss";
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Export the workbook to CSV; formatted date strings will be written to the file
        workbook.Save("output.csv", SaveFormat.CSV);
    }
}
