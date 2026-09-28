// Title: How to add and read a comment on an Aspose.Cells ListObject (Excel table) in C#
// AI Prompts: Create an Excel table with Aspose.Cells, assign a descriptive comment to the ListObject, and output the comment text to the console using C#. | Show how to set the Comment property of a ListObject and then retrieve it programmatically in a .NET workbook.
// Common Searches: aspnet add comment to Excel table using Aspose.Cells ListObject | retrieve ListObject comment Aspose.Cells C# | how to set ListObject.Comment property in Aspose.Cells .NET | read table comment from workbook with Aspose.Cells C# | example of adding metadata comment to Excel table via Aspose.Cells
// Tags: Aspose.Cells ListObject comment | C# set ListObject.Comment | read Excel table comment Aspose.Cells | Aspose.Cells add table metadata | Excel table comment .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// Creates a workbook, adds a ListObject named "Employees", sets a comment describing the table, retrieves the comment text, prints it, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range that will become the table (including header)
            int firstRow = 0;          // zero‑based index
            int firstColumn = 0;
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;

            // Add a ListObject (Excel table) to the worksheet; specify that the range has headers
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a meaningful display name to the table
            table.DisplayName = "Employees";

            // Add a comment describing the purpose of the table
            table.Comment = "This table stores employee IDs and names.";

            // Retrieve the comment text programmatically
            string commentText = table.Comment;

            // Display the comment
            Console.WriteLine("Table comment: " + commentText);

            // Save the workbook (optional)
            string outputPath = "TableWithComment.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
