// Title: Add a ListObject table with filter buttons to range A1:C10 and apply a built‑in style using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to create a ListObject covering cells A1:C10, display the header row, and enable the auto‑filter dropdowns. | Show how to set a built‑in table style (e.g., TableStyleMedium9) for the ListObject and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# create Excel table from specific range with filter dropdowns | How to add a ListObject with header filters to A1:C10 using Aspose.Cells .NET | Apply built‑in table style to an Aspose.Cells ListObject and save workbook
// Tags: Aspose.Cells ListObject creation from cell range | enable auto‑filter on Excel table using Aspose.Cells | apply built‑in table style in Aspose.Cells .NET | save workbook with table to .xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds a ListObject covering cells A1:C10 on the first worksheet, shows the header row to activate filter buttons, applies a built‑in table style, and saves the file as output.xlsx.
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

            // Define the range for the table (A1:C10)
            CellArea tableArea = CellArea.CreateCellArea("A1", "C10");

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                tableArea.StartRow,
                tableArea.StartColumn,
                tableArea.EndRow,
                tableArea.EndColumn,
                true); // true indicates that the first row contains headers

            // Retrieve the created ListObject
            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is shown
            table.ShowHeaderRow = true;

            // (Optional) Apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
