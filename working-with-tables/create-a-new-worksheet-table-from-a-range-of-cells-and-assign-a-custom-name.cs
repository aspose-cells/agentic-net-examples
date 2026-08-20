// Title: Create a Table from a Cell Range and Set a Custom Name with Aspose.Cells for .NET
// Description: Demonstrates how to generate a new workbook, populate cells A1:C4, convert the range into a ListObject table, assign a custom DisplayName (e.g., EmployeeTable), apply a table style, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells create table from range | Aspose.Cells ListObject DisplayName | custom table name Aspose.Cells | apply table style Aspose.Cells | save workbook with table Aspose | C# Aspose.Cells table example
// Common Searches: how to add a table from a range in Aspose.Cells | set custom display name for ListObject Aspose.Cells .NET | Aspose.Cells table style options | save workbook with named table Aspose.Cells | C# create table and assign name using Aspose.Cells
// Developer Intent: Create a worksheet table from a defined range and give it a custom display name.
// Use Cases: Generate a named table for employee data to enable easy reference in formulas and filters. | Apply a predefined table style for consistent visual formatting before exporting reports. | Create multiple named tables on a single sheet to separate distinct datasets such as sales, inventory, and contacts.
// AI Prompts: Write C# code with Aspose.Cells that creates a table from range B2:E10, includes headers, and sets its DisplayName to 'SalesData'. | Show how to change the TableStyleName of an existing ListObject and save the workbook as 'Report.xlsx' using Aspose.Cells. | Provide an example that adds several tables with custom DisplayName values to one worksheet and then retrieves a table by its name.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to generate a new workbook, populate cells A1:C4, convert the range into a ListObject table, assign a custom DisplayName (e.g., EmployeeTable), apply a table style, and save the file as an XLSX document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill sample data in the range A1:C4 (including headers)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");

        int id = 1;
        for (int row = 2; row <= 4; row++)
        {
            cells[row - 1, 0].PutValue(id++);                 // ID column
            cells[row - 1, 1].PutValue("Person" + row);       // Name column
            cells[row - 1, 2].PutValue(50 + row * 5);         // Score column
        }

        // Create a table (ListObject) from the range A1:C4 and assign a custom name
        ListObjectCollection tables = sheet.ListObjects;
        int tableIndex = tables.Add("A1", "C4", true); // hasHeaders = true
        ListObject table = tables[tableIndex];
        table.DisplayName = "EmployeeTable";          // custom table name
        table.TableStyleName = "TableStyleMedium9";   // optional styling

        // Save the workbook
        workbook.Save("TableWithCustomName.xlsx");
    }
}
