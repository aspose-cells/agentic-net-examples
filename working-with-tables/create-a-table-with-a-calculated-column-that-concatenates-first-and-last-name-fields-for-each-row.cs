// Title: Add a calculated FullName column to an Excel table that concatenates FirstName and LastName using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject and sets a formula to join FirstName and LastName into a FullName column. | Show how to use structured references in an Aspose.Cells table to define a derived column that combines two text fields. | Provide a complete example that auto‑fits columns and saves the workbook after adding the calculated column.
// Common Searches: aspnet aspose.cells create table with derived column concatenating names | c# structured reference formula for full name column in Aspose.Cells | how to set ListObject column formula in Aspose.Cells .NET | auto fit columns and save workbook using Aspose.Cells C# example
// Tags: Aspose.Cells column formula with structured references | C# add ListObject to worksheet | concatenate text fields in Excel table using Aspose.Cells | auto‑fit columns Aspose.Cells workbook | save workbook as .xlsx C# Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The program creates a new workbook, inserts a ListObject named PeopleTable with FirstName, LastName, and FullName columns, fills sample data, assigns a structured‑reference formula to the FullName column that joins the first and last names, auto‑fits the columns, and saves the file as PeopleTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers
            sheet.Cells["A1"].PutValue("FirstName");
            sheet.Cells["B1"].PutValue("LastName");
            sheet.Cells["C1"].PutValue("FullName");

            // Sample data
            string[,] data = {
                { "John", "Doe" },
                { "Jane", "Smith" },
                { "Bob", "Johnson" }
            };

            // Populate first and last name columns
            for (int i = 0; i < data.GetLength(0); i++)
            {
                sheet.Cells[i + 1, 0].PutValue(data[i, 0]); // Column A
                sheet.Cells[i + 1, 1].PutValue(data[i, 1]); // Column B
            }

            // Define the range for the table (including header row)
            int totalRows = data.GetLength(0) + 1; // header + data rows
            int totalCols = 3; // FirstName, LastName, FullName

            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = totalRows - 1,      // zero‑based index
                EndColumn = totalCols - 1    // zero‑based index
            };

            // Add a ListObject (Excel table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "PeopleTable";

            // Set formula for the FullName column using structured references
            // Column index 2 corresponds to the third column (FullName)
            table.ListColumns[2].Formula = "=[@FirstName] & \" \" & [@LastName]";

            // Auto‑fit columns for better visibility
            sheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "PeopleTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
