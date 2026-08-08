// Title: Apply Light 1 Slicer Style to a Pivot Table with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Fruit" field, and set the slicer's StyleType to SlicerStyleLight1 using Aspose.Cells for .NET. The workbook is saved as SlicerStyleLight1.xlsx, showcasing quick visual formatting of slicers.
// Keywords: Aspose.Cells slicer style | C# Light1 slicer | SlicerStyleLight1 example | pivot table slicer formatting .NET | apply built‑in slicer style | Excel slicer styling code | Aspose.Cells C# tutorial | GitHub Aspose.Cells slicer sample
// Common Searches: how to set slicer style in Aspose.Cells C# | apply Light 1 style to Excel slicer programmatically | Aspose.Cells pivot table slicer example | C# code for slicer formatting with Aspose.Cells | built‑in slicer styles .NET
// Developer Intent: Programmatically assign the Light 1 built‑in style to a slicer linked to a pivot table.
// Use Cases: Standardize dashboard appearance by applying corporate Light 1 slicer theme. | Generate automated reports with consistently styled slicers for end‑user filtering. | Create teaching material that illustrates slicer styling options in Aspose.Cells.
// AI Prompts: Show code to apply other built‑in slicer styles (e.g., Light2, Dark1) with Aspose.Cells for .NET. | Explain how to change a slicer's style after loading an existing workbook. | List all values of the SlicerStyleType enum and suggest suitable scenarios for each.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    // This example demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Fruit" field, and set the slicer's StyleType to SlicerStyleLight1 using Aspose.Cells for .NET. The workbook is saved as SlicerStyleLight1.xlsx, showcasing quick visual formatting of slicers.
    class ApplySlicerStyle
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully: SlicerStyleLight1.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the slicer source
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(20);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit field as row
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity as data

            // Add a slicer linked to the "Fruit" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "E3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply the built‑in Light 1 slicer style
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1.xlsx");
        }
    }
}
