// Title: Add vertical slicers for each table column with equal spacing using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to load an Excel workbook with Aspose.Cells, ensure a ListObject exists, and programmatically create a slicer for every table column. Each slicer receives a uniform width, height, and caption, then is positioned in a single column with configurable top offset, left offset, and consistent pixel spacing before the workbook is saved.
// Keywords: Aspose.Cells slicer C# | create slicers programmatically | vertical slicer alignment .NET | slicer pixel size Aspose.Cells | add slicer for each table column | Excel slicer spacing | ListObject slicer automation | Aspose.Cells dashboard slicers | C# Excel slicer example
// Common Searches: how to add slicers for all columns of a table using Aspose.Cells | vertical alignment of multiple slicers in C# | set slicer width height pixels Aspose.Cells | programmatically create slicers in Excel with .NET | Aspose.Cells slicer spacing configuration
// Developer Intent: Generate a slicer for every column of a worksheet table and arrange them vertically with equal gaps.
// Use Cases: Build an interactive Excel dashboard where each column filter appears as a neatly stacked slicer. | Automate report templates that need predefined slicers for Category, Product, and Amount fields. | Provide non‑technical users with ready‑to‑use slicers that maintain consistent size and spacing across workbooks.
// AI Prompts: Show how to modify the code to place slicers horizontally with equal spacing instead of vertically. | Add styling to each slicer (background color, font, border) after positioning them. | Explain how to retrieve the created slicer objects to attach custom event handlers or further customize their properties.

using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// This example demonstrates how to load an Excel workbook with Aspose.Cells, ensure a ListObject exists, and programmatically create a slicer for every table column. Each slicer receives a uniform width, height, and caption, then is positioned in a single column with configurable top offset, left offset, and consistent pixel spacing before the workbook is saved.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one table; create a sample table if none exists
        if (worksheet.ListObjects.Count == 0)
        {
            // Sample data for a table (A1:C5)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Product");
            worksheet.Cells["C1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue("Item1");
            worksheet.Cells["C2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue("Item2");
            worksheet.Cells["C3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("A");
            worksheet.Cells["B4"].PutValue("Item3");
            worksheet.Cells["C4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("B");
            worksheet.Cells["B5"].PutValue("Item4");
            worksheet.Cells["C5"].PutValue(40);

            int tableIdx = worksheet.ListObjects.Add("A1", "C5", true);
            worksheet.ListObjects[tableIdx].TableStyleType = TableStyleType.TableStyleMedium2;
        }

        // Get the first table on the worksheet
        ListObject table = worksheet.ListObjects[0];
        SlicerCollection slicers = worksheet.Slicers;

        // Configuration for slicer placement
        int startColumn = 5;          // Column index where slicers will start (e.g., column F)
        int spacingPixels = 10;       // Space between slicers
        int leftPixel = 20;           // Align all slicers to the same left position
        int topStartPixel = 20;       // Top position for the first slicer
        int slicerWidthPixel = 150;   // Uniform width
        int slicerHeightPixel = 120;  // Uniform height

        // Add a slicer for each column in the table and align them vertically
        for (int i = 0; i < table.ListColumns.Count; i++)
        {
            // Add slicer at the same cell location; we'll adjust pixel positions later
            int slicerIdx = slicers.Add(table, table.ListColumns[i], 0, startColumn);
            Slicer slicer = slicers[slicerIdx];

            // Set size
            slicer.WidthPixel = slicerWidthPixel;
            slicer.HeightPixel = slicerHeightPixel;

            // Calculate vertical position
            slicer.TopPixel = topStartPixel + i * (slicerHeightPixel + spacingPixels);
            slicer.LeftPixel = leftPixel;

            // Optional: set a caption matching the column name
            slicer.Caption = table.ListColumns[i].Name;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
