// Title: Create a worksheet table from a cell range and assign a custom display name with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to convert the range A1:C5 into a ListObject, enable the header row, and set its DisplayName to a custom identifier. | Show how to define a CellArea, add a styled table to a worksheet, apply a built‑in table style such as TableStyleMedium9, and save the workbook as an XLSX file using Aspose.Cells. | Provide a complete example that populates sample data, creates a table with a custom display name, applies a table style, and writes the file to disk in C#.
// Common Searches: Aspose.Cells C# create table from range A1:C5 with custom display name | how to add a ListObject with headers and set DisplayName in Aspose.Cells for .NET | apply TableStyleMedium9 to a worksheet table using Aspose.Cells C# | save an Excel workbook containing a styled table to XLSX with Aspose.Cells | define CellArea for table creation in Aspose.Cells C# example
// Tags: worksheet table creation from cell range Aspose.Cells | custom display name for ListObject Aspose.Cells | apply built-in table style Aspose.Cells C# | CellArea definition for table creation Aspose.Cells | save workbook as XLSX with styled table Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample program creates a new workbook, fills cells A1:C5 with sample data, defines a CellArea covering that range, adds a ListObject (table) with headers, assigns a custom DisplayName, applies the TableStyleMedium9 style, and saves the result as CreatedTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (or add a new one)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Fill sample data into cells A1:C5
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Diana");
            sheet.Cells["C5"].PutValue(88);

            // Define the range that will become a table (A1:C5)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 3;
            CellArea tableArea = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstColumn,
                EndRow = firstRow + totalRows - 1,
                EndColumn = firstColumn + totalColumns - 1
            };

            // Add a ListObject (table) to the worksheet based on the defined range
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true); // true indicates that the first row contains column headers

            // Retrieve the created table
            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a custom display name to the table (Name property is not available in this API version)
            table.DisplayName = "MyCustomTable";

            // Optionally, apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook to a file
            string outputPath = "CreatedTable.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
