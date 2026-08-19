// Title: Aspose.Cells for .NET – Set Slicer Width to 150 px and Verify It (C# Example)
// Description: This C# sample creates a workbook, builds a pivot table, adds a slicer linked to the "Fruit" field, sets the slicer width to 150 pixels via the Shape.Width property, prints the actual width to the console for verification, and saves the file as SlicerWidthDemo.xlsx.
// Keywords: Aspose.Cells | C# | slicer width | Shape.Width | 150 pixels | pivot table slicer | Excel automation | verify slicer size | set slicer dimensions | Aspose.Cells example | GitHub
// Common Searches: Aspose.Cells set slicer width C# | How to change slicer size in Excel using Aspose.Cells | C# code to set slicer width to 150 pixels | Verify slicer dimensions programmatically Aspose.Cells | Add slicer to pivot table Aspose.Cells
// Developer Intent: Programmatically set a slicer’s width to 150 px and confirm the value in C# using Aspose.Cells.
// Use Cases: Ensure consistent slicer sizing across automated Excel dashboards. | Validate slicer layout before distributing generated reports. | Apply a fixed slicer width when creating multiple workbooks programmatically. | Create a reusable template for pivot‑table slicers in .NET applications.
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds a pivot table, inserts a slicer for the 'Fruit' field, sets its Shape.Width to 150 pixels, prints the width, and saves the file. | Explain how to read the Shape.Width property of a slicer after setting it with Aspose.Cells for .NET. | Provide step‑by‑step instructions to adjust and verify slicer dimensions in an Excel file using Aspose.Cells C# API.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerWidthDemo
{
    // This C# sample creates a workbook, builds a pivot table, adds a slicer linked to the "Fruit" field, sets the slicer width to 150 pixels via the Shape.Width property, prints the actual width to the console for verification, and saves the file as SlicerWidthDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                sheet.Cells["A1"].PutValue("Fruit");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(90);

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a slicer linked to the pivot table for the "Fruit" field
                // Correct parameter order: destination cell, then field name
                int slicerIdx = sheet.Slicers.Add(pivot, "A6", "Fruit");
                Slicer slicer = sheet.Slicers[slicerIdx];
                slicer.AddPivotConnection(pivot); // ensure connection

                // Set the slicer width to 150 pixels using the Shape object (preferred)
                slicer.Shape.Width = 150;

                // Verify the width by reading back the property and printing it
                Console.WriteLine($"Slicer width (pixels) set to: {slicer.Shape.Width}");

                // Save the workbook to a file
                workbook.Save("SlicerWidthDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
