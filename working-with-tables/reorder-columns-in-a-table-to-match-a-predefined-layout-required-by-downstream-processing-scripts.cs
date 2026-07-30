// Title: Reorder Excel Table Columns with Aspose.Cells CopyColumns in C#
// Description: This example creates a workbook, defines a source ListObject with columns ID, Name, Age, and Salary, and generates a new worksheet where the columns are copied in the order Name, ID, Salary, Age using the CopyColumns method. The destination table’s column headers are refreshed, formatting is preserved, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | CopyColumns | ListObject | reorder columns | Excel table column order | move columns between worksheets | preserve formatting | data transformation | global | United States | Europe
// Common Searches: how to change column order of a ListObject using Aspose.Cells | copy specific columns to another sheet in C# Aspose.Cells | reorder Excel table columns programmatically .NET | preserve table style when moving columns Aspose.Cells | sample code for CopyColumns Aspose.Cells C#
// Developer Intent: The developer needs to rearrange the columns of an Excel table to a predefined sequence without altering the original sheet.
// Use Cases: Create a downstream‑ready worksheet with a custom column layout. | Generate reports where the column sequence differs from the source data. | Prepare data files for systems that require a specific column order.
// AI Prompts: Show how to reorder columns of an existing ListObject in place without creating a new worksheet using Aspose.Cells. | Provide a dynamic method that maps header names to indexes and reorders them with CopyColumns. | Explain how to keep cell formats, table styles, and formulas intact while reordering columns in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsColumnReorderDemo
{
    // This example creates a workbook, defines a source ListObject with columns ID, Name, Age, and Salary, and generates a new worksheet where the columns are copied in the order Name, ID, Salary, Age using the CopyColumns method. The destination table’s column headers are refreshed, formatting is preserved, and the file is saved as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (source)
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate sample data with headers in columns A‑D
            // Headers: ID, Name, Age, Salary
            srcSheet.Cells["A1"].PutValue("ID");
            srcSheet.Cells["B1"].PutValue("Name");
            srcSheet.Cells["C1"].PutValue("Age");
            srcSheet.Cells["D1"].PutValue("Salary");

            // Add some rows
            for (int i = 2; i <= 6; i++)
            {
                srcSheet.Cells[i - 1, 0].PutValue(i - 1);                 // ID
                srcSheet.Cells[i - 1, 1].PutValue($"Person {i - 1}");   // Name
                srcSheet.Cells[i - 1, 2].PutValue(20 + i);              // Age
                srcSheet.Cells[i - 1, 3].PutValue(3000 + i * 100);      // Salary
            }

            // Create a table (ListObject) covering the data range (A1:D6)
            ListObjectCollection srcTables = srcSheet.ListObjects;
            int srcTableIndex = srcTables.Add(0, 0, 5, 3, true);
            ListObject srcTable = srcTables[srcTableIndex];
            srcTable.DisplayName = "EmployeeTable";

            // Desired column order: Name, ID, Salary, Age
            // Corresponding source column indexes (0‑based): B(1), A(0), D(3), C(2)
            int[] desiredOrder = new int[] { 1, 0, 3, 2 };

            // Add a new worksheet that will hold the reordered columns
            Worksheet destSheet = workbook.Worksheets.Add("Reordered");

            // Copy columns from source to destination according to the desired order
            // Each iteration copies one whole column (including header) to the target position
            for (int destCol = 0; destCol < desiredOrder.Length; destCol++)
            {
                int srcCol = desiredOrder[destCol];
                // Copy the whole column (including formats) from srcSheet to destSheet
                destSheet.Cells.CopyColumns(
                    srcSheet.Cells,   // source cells
                    srcCol,           // source column index
                    destCol,          // destination column index in destSheet
                    1);               // number of columns to copy
            }

            // Create a table in the destination sheet that matches the copied range
            ListObjectCollection destTables = destSheet.ListObjects;
            int destTableIndex = destTables.Add(0, 0, 5, desiredOrder.Length - 1, true);
            ListObject destTable = destTables[destTableIndex];
            destTable.DisplayName = "ReorderedEmployeeTable";

            // Update the table column names to reflect the header values in the new layout
            destTable.UpdateColumnName();

            // Save the workbook (output file)
            workbook.Save("ReorderedColumnsDemo.xlsx");
        }
    }
}
