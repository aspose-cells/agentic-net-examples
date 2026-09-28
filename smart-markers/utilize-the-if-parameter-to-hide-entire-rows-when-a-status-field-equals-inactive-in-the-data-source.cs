// Title: How to hide entire rows in an Excel file using Aspose.Cells C# when the Status column equals "Inactive" (If smart‑marker example)
// AI Prompts: Generate C# code with Aspose.Cells that iterates through a worksheet, checks column B for the value "Inactive", and calls Cells.HideRow to hide those rows before saving the workbook. | Demonstrate how to use the Aspose.Cells smart‑marker If parameter to automatically hide rows whose Status field is "Inactive" in a generated Excel report. | Provide a snippet that records the row numbers hidden by the conditional logic, prints them to the console, and then saves the file as an .xlsx document.
// Common Searches: aspnet hide rows where column value is Inactive using Aspose.Cells | c# Aspose.Cells conditional row visibility based on status column | smart markers If parameter hide rows Aspose.Cells example | how to programmatically hide rows in Excel with Aspose.Cells .NET | filter out inactive rows in generated Excel using Aspose.Cells C#
// Tags: conditional row hiding Aspose.Cells C# | hide rows based on column value Excel .NET | smart markers If parameter Aspose.Cells | programmatic Excel row visibility C# | Aspose.Cells hide row API

using System;
using Aspose.Cells;

namespace AsposeCellsRowHideDemo
{
    // The example creates a workbook, fills it with Name and Status data, loops through each data row, and uses Cells.HideRow to hide any row where the Status column equals "Inactive". It logs which rows were hidden and saves the result as RowsHiddenBasedOnStatus.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Status");

            // Sample data rows
            string[,] data = new string[,]
            {
                { "Alice",   "Active"   },
                { "Bob",     "Inactive" },
                { "Charlie", "Active"   },
                { "Diana",   "Inactive" },
                { "Eve",     "Active"   }
            };

            // Populate the worksheet with sample data
            for (int i = 0; i < data.GetLength(0); i++)
            {
                int rowIndex = i + 1; // +1 to account for header row
                cells[$"A{rowIndex + 1}"].PutValue(data[i, 0]); // Name column (A)
                cells[$"B{rowIndex + 1}"].PutValue(data[i, 1]); // Status column (B)
            }

            // Hide rows where the Status column equals "Inactive"
            // Status column is column index 1 (B)
            int firstDataRow = 1; // zero‑based index of first data row (row 2 in Excel)
            int lastDataRow = cells.MaxDataRow; // last row with data

            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                // Retrieve the status value as string
                string status = cells[row, 1].StringValue; // column B

                // If status is "Inactive", hide the entire row
                if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                {
                    cells.HideRow(row);
                }
            }

            // Optional: Verify hidden rows (for debugging)
            for (int row = 0; row <= lastDataRow; row++)
            {
                bool hidden = cells.IsRowHidden(row);
                Console.WriteLine($"Row {row + 1} hidden: {hidden}");
            }

            // Save the workbook
            workbook.Save("RowsHiddenBasedOnStatus.xlsx", SaveFormat.Xlsx);
        }
    }
}
