// Title: Conditionally hide Excel columns in C# with Aspose.Cells using a DataTable visibility flag
// Common Searches: Aspose.Cells hide column when DataTable IsVisible is false in C# | C# code to conditionally hide Excel columns using Aspose.Cells HideColumn | How to control column visibility in an Aspose.Cells workbook with a boolean flag | Using UnhideColumn with default width after conditional hide in Aspose.Cells | Smart markers If parameter to hide columns based on data source in C#
// Tags: HideColumn method Aspose.Cells C# | conditional column visibility Aspose.Cells | DataTable driven column hiding Aspose.Cells | UnhideColumn default width Aspose.Cells | smart markers If parameter column visibility

using System;
using System.Data;
using Aspose.Cells;

namespace HideColumnsConditionalDemo
{
    // The example creates a workbook, fills three columns with sample data, defines a DataTable containing ColumnIndex and IsVisible flags, iterates through the table to hide columns where IsVisible is false using sheet.Cells.HideColumn and to unhide others with a default width via sheet.Cells.UnhideColumn, and finally saves the workbook as ConditionalHiddenColumns.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare a sample workbook with some data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill the worksheet with sample data (3 columns, 5 rows)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);                     // ID
                    sheet.Cells[$"B{i}"].PutValue($"Item {i - 1}");          // Name
                    sheet.Cells[$"C{i}"].PutValue(10 * (i - 1));            // Score
                }

                // -----------------------------------------------------------------
                // 2. Create a data source that defines which columns should be visible.
                //    The data source contains two fields:
                //      - ColumnIndex : zero‑based index of the column in the worksheet.
                //      - IsVisible   : flag indicating whether the column must be shown.
                // -----------------------------------------------------------------
                DataTable visibilityTable = new DataTable("ColumnVisibility");
                visibilityTable.Columns.Add("ColumnIndex", typeof(int));
                visibilityTable.Columns.Add("IsVisible", typeof(bool));

                // Define visibility: show column 0 (ID), hide column 1 (Name), show column 2 (Score)
                visibilityTable.Rows.Add(0, true);   // ID column – visible
                visibilityTable.Rows.Add(1, false);  // Name column – hidden
                visibilityTable.Rows.Add(2, true);   // Score column – visible

                // -----------------------------------------------------------------
                // 3. Use the visibility table to hide or unhide columns.
                // -----------------------------------------------------------------
                foreach (DataRow row in visibilityTable.Rows)
                {
                    int columnIndex = Convert.ToInt32(row["ColumnIndex"]);
                    bool isVisible = Convert.ToBoolean(row["IsVisible"]);

                    if (!isVisible)
                    {
                        // Hide the column.
                        sheet.Cells.HideColumn(columnIndex);
                        Console.WriteLine($"Column at index {columnIndex} has been hidden because IsVisible = false.");
                    }
                    else
                    {
                        // Ensure the column is visible (unhide) with a default width.
                        sheet.Cells.UnhideColumn(columnIndex, 10.0);
                        Console.WriteLine($"Column at index {columnIndex} is kept visible.");
                    }
                }

                // -----------------------------------------------------------------
                // 4. Save the modified workbook to verify the result.
                // -----------------------------------------------------------------
                string outputPath = "ConditionalHiddenColumns.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'. Open it to see the hidden column(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
