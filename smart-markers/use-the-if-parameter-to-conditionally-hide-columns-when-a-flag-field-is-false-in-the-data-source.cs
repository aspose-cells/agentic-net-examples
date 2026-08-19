// Title: Conditionally hide Excel columns with Aspose.Cells for .NET using an If‑parameter flag
// Description: This C# example demonstrates how to create a workbook, fill it with sample data, and control column visibility with a DataTable that stores a boolean flag for each column. Columns whose flag is false are hidden using Cells.HideColumn, while true flags keep the column visible with Cells.UnhideColumn. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells hide column C# | conditional column visibility .NET | Excel column hide based on flag | smart markers If parameter Aspose | Cells.HideColumn example | dynamic column visibility Aspose.Cells
// Common Searches: how to hide Excel columns programmatically with Aspose.Cells | Aspose.Cells conditional column visibility using smart markers | C# hide column if flag is false Aspose | use DataTable to control column visibility in Aspose.Cells | Aspose.Cells hide column based on boolean field
// Developer Intent: The developer needs to show or hide worksheet columns at runtime according to a boolean flag supplied by a data source.
// Use Cases: Generate role‑based reports where columns are hidden for users without permission. | Export data sets that omit optional fields when a configuration flag is false. | Create reusable Excel templates that automatically collapse columns based on source‑data flags.
// AI Prompts: Show how to use Aspose.Cells smart markers If parameter to hide a column when a flag field is false. | Provide C# code that reads a DataTable of column visibility flags and applies HideColumn/UnhideColumn in Aspose.Cells. | Explain how to combine Aspose.Cells conditional logic with column hiding based on a boolean column in the data source.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook, fill it with sample data, and control column visibility with a DataTable that stores a boolean flag for each column. Columns whose flag is false are hidden using Cells.HideColumn, while true flags keep the column visible with Cells.UnhideColumn. The workbook is then saved as an .xlsx file.
    public class ConditionalColumnHideDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in columns A to D
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Score");
                cells["D1"].PutValue("Remarks");

                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Alice");
                cells["C2"].PutValue(85);
                cells["D2"].PutValue("Good");

                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Bob");
                cells["C3"].PutValue(92);
                cells["D3"].PutValue("Excellent");

                // Simulate a flag data source that determines visibility of each column
                // ColumnIndex is zero‑based (0 = A, 1 = B, etc.)
                DataTable flagTable = new DataTable();
                flagTable.Columns.Add("ColumnIndex", typeof(int));
                flagTable.Columns.Add("IsVisible", typeof(bool));

                // Define visibility: hide column B (index 1) and column D (index 3)
                flagTable.Rows.Add(0, true);   // Column A visible
                flagTable.Rows.Add(1, false);  // Column B hidden
                flagTable.Rows.Add(2, true);   // Column C visible
                flagTable.Rows.Add(3, false);  // Column D hidden

                // Iterate through the flag table and hide columns where IsVisible is false
                foreach (DataRow row in flagTable.Rows)
                {
                    int colIndex = (int)row["ColumnIndex"];
                    bool isVisible = (bool)row["IsVisible"];

                    if (!isVisible)
                    {
                        // Hide the column
                        cells.HideColumn(colIndex);
                    }
                    else
                    {
                        // Ensure the column is visible
                        cells.UnhideColumn(colIndex, cells.StandardWidth);
                    }
                }

                // Save the workbook to a file
                workbook.Save("ConditionalColumnHideDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConditionalColumnHideDemo.Run();
        }
    }
}
