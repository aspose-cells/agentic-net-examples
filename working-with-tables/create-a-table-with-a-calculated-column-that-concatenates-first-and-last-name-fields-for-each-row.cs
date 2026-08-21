// Title: Add a Calculated FullName Column to an Excel Table with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, defines a ListObject with FirstName and LastName fields, expands the table, adds a FullName column, and applies a structured‑reference formula (=[@FirstName] & " " & [@LastName]) so each row automatically concatenates the two name values. The file is saved as an .xlsx document.
// Keywords: Aspose.Cells | C# | Excel table | ListObject | calculated column | concatenate columns | FullName formula | structured reference | resize table | add column programmatically
// Common Searches: Aspose.Cells add calculated column | C# concatenate first and last name in Excel table | Resize ListObject and set formula Aspose | Structured reference formula Aspose.Cells | Create FullName column programmatically
// Developer Intent: Programmatically insert a new column into an Aspose.Cells ListObject and set a formula that joins FirstName and LastName into a FullName value.
// Use Cases: Generate contact sheets where the full name updates automatically when source fields change. | Build employee export files with a computed FullName column without manual data entry. | Create dynamic reports that keep name concatenations in sync across large datasets.
// AI Prompts: Show C# code using Aspose.Cells to add a FullName column to an existing Excel table and apply a concatenation formula. | Explain how to resize an Aspose.Cells ListObject and assign a structured‑reference expression for merging two text columns. | Provide a step‑by‑step guide for creating a calculated column that combines FirstName and LastName with a space separator.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, defines a ListObject with FirstName and LastName fields, expands the table, adds a FullName column, and applies a structured‑reference formula (=[@FirstName] & " " & [@LastName]) so each row automatically concatenates the two name values. The file is saved as an .xlsx document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: FirstName and LastName columns
            cells["A1"].PutValue("FirstName");
            cells["B1"].PutValue("LastName");
            cells["A2"].PutValue("John");
            cells["B2"].PutValue("Doe");
            cells["A3"].PutValue("Jane");
            cells["B3"].PutValue("Smith");

            // Add a ListObject (Excel table) covering the data range A1:B3
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "People";
            table.ShowHeaderRow = true;
            table.ShowTotals = false;

            // Determine current size of the table
            int rowCount = table.DataRange.RowCount;          // data rows (excluding header)
            int columnCount = table.DataRange.ColumnCount;    // existing data columns

            // Expand the table to include a new column for the calculated FullName
            // Resize requires the hasHeaders flag; we keep it true because the table has a header row
            table.Resize(table.StartRow, table.StartColumn, rowCount, columnCount + 1, true);

            // Access the newly added column (last column in the table)
            ListColumn fullNameColumn = table.ListColumns[table.ListColumns.Count - 1];
            fullNameColumn.Name = "FullName";

            // Set the calculated column formula to concatenate FirstName and LastName
            // Structured reference syntax: =[@FirstName] & " " & [@LastName]
            fullNameColumn.Formula = "=[@FirstName] & \" \" & [@LastName]";

            // Save the workbook to a file
            string outputPath = "PeopleTable.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
