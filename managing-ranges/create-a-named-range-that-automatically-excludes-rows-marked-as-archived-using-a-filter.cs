// Title: Aspose.Cells C# – Create a Named Range that Skips “Archived” Rows Using AutoFilter
// Description: This example builds a workbook, adds ID, Name and Status columns, applies an AutoFilter that hides rows where Status = "Archived", defines a named range "ActiveData" that points to the filtered area, and saves the file as NamedRangeWithFilter.xlsx.
// Keywords: Aspose.Cells | C# | named range | AutoFilter | exclude archived rows | dynamic range | filter NotEqual | Excel named range with filter | Aspose.Cells example | programmatic row filter
// Common Searches: Aspose.Cells filter out rows with specific value | Create named range that ignores archived rows C# | AutoFilter NotEqual operator Aspose.Cells example | Dynamic named range based on filter Aspose.Cells | C# code to hide archived rows and define named range
// Developer Intent: Generate a named range that contains only rows where the Status column is not "Archived" by using an AutoFilter in Aspose.Cells.
// Use Cases: Build reports that automatically exclude archived records. | Copy or export only active rows to another sheet or workbook. | Reference the filtered range in formulas, charts, or pivot tables without manual updates.
// AI Prompts: Write C# code with Aspose.Cells to create a named range that excludes rows where Status = 'Archived' using AutoFilter. | Show how to change the filter criteria to include only rows where Status = 'Active' and update the named range accordingly. | Explain how to use the "ActiveData" named range in a chart or formula after the filter is applied.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeWithFilter
{
    // This example builds a workbook, adds ID, Name and Status columns, applies an AutoFilter that hides rows where Status = "Archived", defines a named range "ActiveData" that points to the filtered area, and saves the file as NamedRangeWithFilter.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate headers
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Status");

            // Sample data (some rows are marked as "Archived")
            string[,] data = new string[,]
            {
                { "1", "Alpha",   "Active"   },
                { "2", "Beta",    "Archived" },
                { "3", "Gamma",   "Active"   },
                { "4", "Delta",   "Archived" },
                { "5", "Epsilon", "Active"   }
            };

            // Fill the worksheet with the sample data
            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 2; // Data starts from row 2 (zero‑based index 1)
                sheet.Cells[row - 1, 0].PutValue(data[i, 0]); // ID
                sheet.Cells[row - 1, 1].PutValue(data[i, 1]); // Name
                sheet.Cells[row - 1, 2].PutValue(data[i, 2]); // Status
            }

            // Determine the last row index (zero‑based)
            int lastRow = sheet.Cells.MaxDataRow;

            // Apply an AutoFilter to the header row covering columns A‑C
            // SetRange(startRow, startColumn, endRow)
            sheet.AutoFilter.SetRange(0, 0, lastRow);

            // Exclude rows where Status = "Archived"
            // Column index for "Status" is 2 (C column)
            // Use a custom filter with NotEqual operator
            sheet.AutoFilter.Custom(2, FilterOperatorType.NotEqual, "Archived");
            sheet.AutoFilter.Refresh();

            // Create a named range that refers to the whole data area (including header)
            // The filter will hide the "Archived" rows, so the named range effectively contains only active rows.
            int nameIndex = workbook.Worksheets.Names.Add("ActiveData");
            Name activeRange = workbook.Worksheets.Names[nameIndex];
            // RefersTo must be a formula string, e.g., "=Sheet1!$A$2:$C$6"
            activeRange.RefersTo = $"=Sheet1!$A$2:$C${lastRow + 1}";

            // Save the workbook
            workbook.Save("NamedRangeWithFilter.xlsx");
        }
    }
}
