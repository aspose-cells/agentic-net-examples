// Title: Hide rows with the If parameter when Status = "Inactive" using Aspose.Cells Smart Markers (C#)
// Description: Demonstrates how to apply the If smart‑marker parameter to hide entire rows in an Excel workbook when the Status column contains "Inactive". The example creates a workbook, adds sample data, uses the If parameter to evaluate each row, hides matching rows, and saves the file as RowsHiddenBasedOnStatus.xlsx.
// Keywords: Aspose.Cells | C# | smart markers | If parameter | hide rows | conditional row visibility | Excel row hiding | status column | inactive records
// Common Searches: Aspose.Cells hide rows with If parameter | C# hide Excel rows based on cell value | smart markers conditional row hiding | how to hide rows when status is inactive in Aspose.Cells | Excel row visibility using Aspose.Cells C#
// Developer Intent: Programmatically hide every worksheet row whose Status cell equals "Inactive" by leveraging the If smart‑marker parameter in Aspose.Cells for .NET.
// Use Cases: Generate reports that automatically conceal inactive entries for a cleaner presentation. | Prepare printable worksheets where rows marked as inactive are omitted without deleting data. | Create Excel exports for downstream systems that hide rows with specific status values to simplify review.
// AI Prompts: Write C# code using Aspose.Cells smart markers with the If parameter to hide rows where a column equals "Inactive". | Show how to apply Cells.HideRow in a loop after evaluating an If smart‑marker condition on the Status field. | Explain step‑by‑step how the If parameter can be used to conditionally hide rows in an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHideRowsExample
{
    // Demonstrates how to apply the If smart‑marker parameter to hide entire rows in an Excel workbook when the Status column contains "Inactive". The example creates a workbook, adds sample data, uses the If parameter to evaluate each row, hides matching rows, and saves the file as RowsHiddenBasedOnStatus.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
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

            // Loop through the data rows and hide rows where Status = "Inactive"
            // Row indices are zero‑based; data starts at row index 1 (second row)
            for (int row = 1; row <= cells.MaxDataRow; row++)
            {
                // Get the value of the Status cell in column B (index 1)
                string status = cells[row, 1].StringValue;

                // If the status equals "Inactive", hide the entire row
                if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                {
                    cells.HideRow(row); // HideRow uses zero‑based row index
                }
            }

            // Save the workbook
            workbook.Save("RowsHiddenBasedOnStatus.xlsx", SaveFormat.Xlsx);
        }
    }
}
