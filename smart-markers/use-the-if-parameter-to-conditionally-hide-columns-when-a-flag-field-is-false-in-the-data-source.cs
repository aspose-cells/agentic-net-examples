using System;
using Aspose.Cells;

namespace AsposeCellsConditionalColumnHide
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data headers (row 0)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Amount");
            cells["D1"].PutValue("Date");

            // Sample data rows (starting at row 2)
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(100);
            cells["D2"].PutValue(DateTime.Today);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(200);
            cells["D3"].PutValue(DateTime.Today.AddDays(1));

            // Flag row (row 1, index 0-based = row 2 in Excel) indicating visibility of each column
            // True = show, False = hide
            cells["A2"].PutValue(true);   // ID column visible
            cells["B2"].PutValue(false);  // Name column hidden
            cells["C2"].PutValue(true);   // Amount column visible
            cells["D2"].PutValue(false);  // Date column hidden

            // Determine the range of columns to evaluate (based on header row)
            int totalColumns = sheet.Cells.MaxColumn + 1; // includes all used columns

            // Iterate through each column and hide it if the flag is false
            for (int col = 0; col < totalColumns; col++)
            {
                // Read the flag value from the flag row (row index 1)
                object flagObj = cells[1, col].Value; // row 1 (second row), column 'col'
                bool isVisible = true; // default to visible

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
                    // Option 1: use HideColumn method
                    cells.HideColumn(col);
                    // Optionally, also set IsHidden property for clarity
                    // sheet.Cells.Columns[col].IsHidden = true;
                }
            }

            // Optionally, remove the flag row from the final output
            // sheet.Cells.DeleteRow(1); // deletes the flag row (row index 1)

            // Save the workbook (lifecycle save)
            workbook.Save("ConditionalColumnHide.xlsx");
        }
    }
}