using System;
using Aspose.Cells;

namespace AsposeCellsConditionalColumnHide
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Row 0: Headers
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            // Row 1: Sample data
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(50);
            cells["C2"].PutValue(1.2);

            // Row 2: Flag fields indicating whether the column should be visible (true = show, false = hide)
            // In a real scenario this row could come from an external data source.
            cells["A3"].PutValue(true);   // Show Product column
            cells["B3"].PutValue(false);  // Hide Quantity column
            cells["C3"].PutValue(true);   // Show Price column

            // Determine the number of columns to evaluate (based on the header row)
            int totalColumns = cells.MaxColumn + 1; // MaxColumn is zero‑based

            // Loop through each column and hide it if the corresponding flag is false
            for (int col = 0; col < totalColumns; col++)
            {
                // Read the flag value from the flag row (row index 2)
                object flagObj = cells[2, col].Value; // Row 2 (zero‑based) holds the flags
                bool isVisible = true; // Default to visible

                if (flagObj != null && flagObj is bool)
                {
                    isVisible = (bool)flagObj;
                }
                else if (flagObj != null && bool.TryParse(flagObj.ToString(), out bool parsed))
                {
                    isVisible = parsed;
                }

                // Hide the column when the flag is false
                if (!isVisible)
                {
                    // Using Cells.HideColumn method to hide the column
                    cells.HideColumn(col);
                }
            }

            // Save the workbook to an Excel file
            workbook.Save("ConditionalColumnHideDemo.xlsx");
        }
    }
}