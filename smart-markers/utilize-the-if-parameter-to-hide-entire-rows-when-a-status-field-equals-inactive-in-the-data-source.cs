// Title: Aspose.Cells C# – Conditional row hiding with Smart Marker If
// Description: Demonstrates creating a workbook, adding ID and Status columns, populating rows, and using Cells.HideRow inside an If check to conceal rows whose Status cell contains "Inactive". The file is saved as RowsHiddenWhenInactive.xlsx.
// Keywords: Aspose.Cells C# hide rows | Excel conditional row visibility | Smart marker If parameter | Hide rows based on cell value | C# Aspose.Cells hide inactive rows | Aspose.Cells US | Aspose.Cells Europe
// Common Searches: Aspose.Cells hide rows based on column value | C# hide Excel rows when status is Inactive | Smart marker If to hide rows Aspose | Programmatically hide rows in .NET Excel using Aspose.Cells | Conditional row hiding vs formatting Aspose.Cells
// Developer Intent: Programmatically conceal rows whose Status column equals "Inactive" in a generated Excel workbook.
// Use Cases: Create a client‑facing report that visually omits deactivated records while keeping them in the file for audit purposes. | Generate a printable spreadsheet that shows only active entries, simplifying stakeholder review. | Build a data‑driven dashboard that automatically collapses rows with a specific status to keep the view focused on current items.
// AI Prompts: Write C# code using Aspose.Cells that hides rows where column B contains "Inactive" by applying the Smart Marker If parameter. | Show an Aspose.Cells example that applies an If condition to hide rows during data population from a DataTable. | Provide a .NET snippet that hides rows based on a status field while preserving row indices for later processing.

using System;
using Aspose.Cells;

namespace AsposeCellsHideRowsExample
{
    // Demonstrates creating a workbook, adding ID and Status columns, populating rows, and using Cells.HideRow inside an If check to conceal rows whose Status cell contains "Inactive". The file is saved as RowsHiddenWhenInactive.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with a "Status" column (column B)
            // Header
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Status");

            // Data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Active");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Inactive");

            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Active");

            cells["A5"].PutValue(4);
            cells["B5"].PutValue("Inactive");

            // Determine the last row with data
            int lastRow = cells.MaxDataRow; // zero‑based index

            // Loop through each data row (skip header at row 0)
            for (int row = 1; row <= lastRow; row++)
            {
                // Get the value of the Status cell (column B, index 1)
                string status = cells[row, 1].StringValue;

                // If the status equals "Inactive", hide the entire row
                if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                {
                    cells.HideRow(row); // row index is zero‑based
                }
            }

            // Save the workbook
            workbook.Save("RowsHiddenWhenInactive.xlsx", SaveFormat.Xlsx);
        }
    }
}
