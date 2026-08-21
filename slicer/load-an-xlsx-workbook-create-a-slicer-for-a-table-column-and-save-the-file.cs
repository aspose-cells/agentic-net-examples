// Title: C# – Add a slicer to a table column in an XLSX workbook with Aspose.Cells
// Description: Shows how to load an existing XLSX file using Aspose.Cells for .NET, ensure a ListObject (Excel table) is present, insert a slicer linked to the first column, set its caption and position, and save the workbook as a new file.
// Keywords: Aspose.Cells | C# slicer | Excel slicer programmatically | ListObject | XLSX | add slicer .NET | worksheet slicer | Excel table slicer | filter UI | Aspose.Cells example
// Common Searches: Aspose.Cells add slicer to Excel table C# | Create slicer for ListObject using Aspose.Cells | C# code to insert slicer at specific cell | How to programmatically add Excel slicer with Aspose | Set slicer caption Aspose.Cells | Load workbook and add slicer .NET
// Developer Intent: Insert a slicer for a table column in an existing workbook and persist the changes.
// Use Cases: Provide interactive filtering in automatically generated reports. | Create a table on‑the‑fly when source data lacks one, then attach a slicer. | Customize slicer appearance—caption, placement, and style—for branded Excel outputs. | Batch‑process multiple workbooks to add consistent slicers across dashboard sheets.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, creates a ListObject if missing, adds a slicer for column 0 at cell D1, sets caption 'Category', and saves as output.xlsx. | Explain step‑by‑step how to position and style a slicer using Aspose.Cells for .NET. | Provide a GitHub‑style snippet that demonstrates adding a slicer to an existing Excel file with error handling.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// Shows how to load an existing XLSX file using Aspose.Cells for .NET, ensure a ListObject (Excel table) is present, insert a slicer linked to the first column, set its caption and position, and save the workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one table (ListObject)
        // If no table exists, create a simple one for demonstration purposes
        if (worksheet.ListObjects.Count == 0)
        {
            // Sample data for the table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("A");
            worksheet.Cells["B4"].PutValue(30);

            // Add a table covering the sample data (rows 0‑3, columns 0‑1)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
        }

        // Retrieve the first table in the worksheet
        ListObject table = worksheet.ListObjects[0];

        // Add a slicer for the first column of the table.
        // The slicer will be placed with its upper‑left corner at cell D1.
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");

        // Optional: customize the slicer (e.g., set a caption)
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Caption = "Category Slicer";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
