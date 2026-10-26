// Title: How to create an Excel table from range A1:C10 with filter buttons using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create a table from A1:C10 and show filter dropdowns | Add a ListObject with headers and filter buttons to a worksheet using Aspose.Cells | Enable Excel table header filters in C# with Aspose.Cells ListObjects | Set ShowHeaderRow property for a table range in Aspose.Cells | Define CellArea A1:C10 and add a ListObject in Aspose.Cells example
// Tags: create ListObject from CellArea Aspose.Cells | enable filter dropdowns on Excel table header Aspose.Cells | show header row property Aspose.Cells C# | define range A1:C10 for Excel table Aspose.Cells | assign display name to ListObject Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a new workbook, defines the CellArea covering A1:C10, adds a ListObject (Excel table) with headers, enables filter dropdowns by setting ShowHeaderRow to true, assigns a display name, and saves the file as TableWithFilters.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:C10 for the table (0‑based indices)
            CellArea tableArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 2
            };

            // Add a ListObject (Excel table) with headers (hasHeaders = true)
            int listObjectIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true); // hasHeaders

            ListObject table = sheet.ListObjects[listObjectIndex];

            // Optional: give the table a display name
            table.DisplayName = "MyTable";

            // Ensure the header row is shown (filter dropdowns appear automatically)
            table.ShowHeaderRow = true;

            // Save the workbook to a file
            string outputPath = "TableWithFilters.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
