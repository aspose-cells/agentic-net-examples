// Title: Aspose.Cells for .NET – Add Group Labels Above Outline Rows (C#)
// Description: Demonstrates how to create a workbook, group rows, and insert free‑floating label shapes positioned just above each group using Placement, Top and Left properties, then save the file as Excel.
// Keywords: Aspose.Cells | C# | group rows | outline rows | label shape | label position | Placement property | free floating label | worksheet shapes | Excel automation
// Common Searches: Aspose.Cells add label above grouped rows | C# place label shape before outline group | set label position in Excel using Aspose.Cells | group rows with headings Aspose.Cells .NET
// Developer Intent: Add heading labels before each outlined row group in an Excel worksheet.
// Use Cases: Generate section headers for financial statements with collapsible row groups | Create dynamic report sections where headings stay visible when rows are collapsed | Export Excel files with custom‑styled group titles independent of cell data
// AI Prompts: Provide C# code that loops through all outline groups and adds a label shape above each, calculating Top and Left values automatically. | Explain the effect of the Placement, Top, and Left properties on label positioning relative to grouped rows in Aspose.Cells. | Show how to style the label (font, background, border) and adjust its size for group headings in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, group rows, and insert free‑floating label shapes positioned just above each group using Placement, Top and Left properties, then save the file as Excel.
class GroupLabelDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for two groups
        // Group 1 (rows 0‑2)
        cells["A1"].PutValue("Group 1 Item 1");
        cells["B1"].PutValue(10);
        cells["A2"].PutValue("Group 1 Item 2");
        cells["B2"].PutValue(20);
        cells["A3"].PutValue("Group 1 Item 3");
        cells["B3"].PutValue(30);

        // Group 2 (rows 3‑5)
        cells["A4"].PutValue("Group 2 Item 1");
        cells["B4"].PutValue(40);
        cells["A5"].PutValue("Group 2 Item 2");
        cells["B5"].PutValue(50);
        cells["A6"].PutValue("Group 2 Item 3");
        cells["B6"].PutValue(60);

        // Create outline groups for the rows
        cells.GroupRows(0, 2, false); // Group 1
        cells.GroupRows(3, 5, false); // Group 2

        // Add a label shape before the first group
        // Parameters: upperLeftRow, upperLeftColumn, top, left, width, height
        Label group1Label = worksheet.Shapes.AddLabel(0, 0, 0, 0, 120, 20);
        group1Label.Text = "Group 1";
        group1Label.Placement = PlacementType.FreeFloating;
        // Position the label (pixels) – placed just above the first data row
        group1Label.Top = 5;
        group1Label.Left = 5;

        // Add a label shape before the second group
        Label group2Label = worksheet.Shapes.AddLabel(3, 0, 0, 0, 120, 20);
        group2Label.Text = "Group 2";
        group2Label.Placement = PlacementType.FreeFloating;
        group2Label.Top = 5;
        group2Label.Left = 5;

        // Save the workbook
        workbook.Save("GroupLabelsDemo.xlsx");
    }
}
