// Title: C# – Enumerate Cells Within a Defined Display Range Using Aspose.Cells
// Description: Learn how to create a visible area with CellArea, build a matching Range object, and iterate only the cells inside that area in Aspose.Cells for .NET. The example shows a workaround for the missing Worksheet.DisplayRange property, prints each cell address and value, and saves the workbook.
// Keywords: Aspose.Cells enumerate cells | C# display range | CellArea to Range conversion | iterate visible cells Aspose.Cells | Worksheet.DisplayRange workaround | Aspose.Cells .NET sample | range enumeration C#
// Common Searches: how to loop through cells in a specific area using Aspose.Cells | Aspose.Cells create range from CellArea | enumerate only visible cells Aspose.Cells .NET | workaround for missing Worksheet.DisplayRange property | C# sample to read cells A1:C3 with Aspose.Cells
// Developer Intent: Iterate over and process only the cells that lie inside a user‑defined display region of a worksheet.
// Use Cases: Generate a report by reading only the cells visible to the user (e.g., A1:C3). | Copy or export a selected worksheet area while ignoring hidden rows and columns. | Apply formatting, formulas, or data validation exclusively to a specific cell block.
// AI Prompts: Provide C# code that uses Aspose.Cells to convert a CellArea (A1:C3) into a Range and enumerate each cell’s address and value. | Show a method that replaces the unavailable Worksheet.DisplayRange property by returning a Range for given start and end addresses. | Explain how to filter cells by a defined visible area when Worksheet.DisplayRange is not supported in the current Aspose.Cells version.

using System;
using System.Collections;
using Aspose.Cells;

// Learn how to create a visible area with CellArea, build a matching Range object, and iterate only the cells inside that area in Aspose.Cells for .NET. The example shows a workaround for the missing Worksheet.DisplayRange property, prints each cell address and value, and saves the workbook.
class DisplayRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue(100);
            cells["B2"].PutValue(200);
            cells["C3"].PutValue(300);
            cells["D4"].PutValue(400);

            // Define the visible (display) range of the worksheet (A1:C3)
            CellArea displayArea = CellArea.CreateCellArea("A1", "C3");

            // NOTE: Worksheet.DisplayRange property is not available in the current Aspose.Cells version.
            // The demo focuses on enumerating cells within the desired area.

            // Build a Range object that matches the display area dimensions
            int totalRows = displayArea.EndRow - displayArea.StartRow + 1;
            int totalCols = displayArea.EndColumn - displayArea.StartColumn + 1;
            Aspose.Cells.Range visibleRange = cells.CreateRange(
                displayArea.StartRow,
                displayArea.StartColumn,
                totalRows,
                totalCols);

            // Enumerate only the cells that lie inside the defined display range
            IEnumerator enumerator = visibleRange.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // Save the workbook (optional, just to verify the result)
            string outputPath = "DisplayRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
