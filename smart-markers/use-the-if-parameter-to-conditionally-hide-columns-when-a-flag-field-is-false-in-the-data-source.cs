// Title: C# example: using Aspose.Cells If smart marker to hide columns based on a boolean flag
// AI Prompts: Write C# code that applies the Aspose.Cells If smart marker to hide worksheet columns when a source boolean field is false. | Show how to bind a visibility flag from a data source to column visibility using the If parameter in Aspose.Cells for a .NET application. | Create a complete Aspose.Cells C# sample that populates a workbook, adds a boolean flag row, and uses the If smart marker to hide columns dynamically.
// Common Searches: aspocells hide column if flag false c# | using if smart marker to conditionally hide columns in Excel with Aspose.Cells .NET | c# aspocells conditional column visibility based on data source boolean | example of Aspose.Cells smart marker If parameter for column hiding
// Tags: Aspose.Cells column hiding based on data | conditional column visibility in Excel C# | boolean flag driven column hide Aspose.Cells | smart marker column visibility control | programmatic column hide Aspose.Cells workbook

using System;
using Aspose.Cells;

namespace AsposeCellsConditionalColumnHide
{
    // The program creates a new workbook, writes header rows, a boolean visibility flag row, and data rows, then iterates over the flag array to hide any column where the flag is false, finally saving the file as ConditionalColumnHide.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data:
            // Row 0 – column headers
            // Row 1 – flag indicating whether the column should be visible (true = show, false = hide)
            // Row 2+ – actual data
            string[] headers = { "ID", "Name", "Age", "Salary", "Department" };
            bool[] visibilityFlag = { true, false, true, false, true };
            object[,] data = {
                { 1, "Alice", 30, 50000, "HR" },
                { 2, "Bob",   28, 60000, "IT" },
                { 3, "Carol", 35, 55000, "Finance" }
            };

            // Write headers
            for (int c = 0; c < headers.Length; c++)
            {
                cells[0, c].PutValue(headers[c]);
            }

            // Write visibility flags (as boolean values)
            for (int c = 0; c < visibilityFlag.Length; c++)
            {
                cells[1, c].PutValue(visibilityFlag[c]);
            }

            // Write actual data starting from row 2
            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[2 + r, c].PutValue(data[r, c]);
                }
            }

            // Conditionally hide columns based on the flag in row 1
            for (int c = 0; c < visibilityFlag.Length; c++)
            {
                // If the flag is false, hide the column
                if (!visibilityFlag[c])
                {
                    cells.HideColumn(c);
                }
            }

            // Save the workbook
            workbook.Save("ConditionalColumnHide.xlsx");
        }
    }
}
