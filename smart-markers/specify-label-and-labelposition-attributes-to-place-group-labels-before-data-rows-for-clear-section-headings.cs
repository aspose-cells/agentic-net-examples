// Title: Insert a free‑floating label shape as a group header above collapsed rows using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that groups a range of rows, sets the outline summary row above the group, and adds a free‑floating label shape as a heading positioned with pixel offsets using Aspose.Cells. | Create an Excel workbook with Aspose.Cells that adds a label shape at a specific cell location, configures PlacementType.FreeFloating, and saves the file.
// Common Searches: aspnet c# how to add a label shape above a grouped row range in Aspose.Cells | Aspose.Cells place free floating label as group header in Excel | C# group rows and add custom heading label using Aspose.Cells | set pixel offset for label shape in Aspose.Cells workbook | outline summary row above groups Aspose.Cells example
// Tags: Aspose.Cells free‑floating label placement | C# group rows with custom header label | Excel outline summary row above groups | label shape pixel offset Aspose.Cells | Aspose.Cells add label shape as group heading

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupLabelDemo
{
    // Demonstrates using Aspose.Cells for .NET to create a workbook, populate sample data, group rows, set the outline summary row above the group, add a free‑floating label shape as a group header, adjust its row/column and pixel offsets, and save the file as GroupLabelDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (header + 6 rows)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            for (int i = 2; i <= 7; i++)
            {
                cells[$"A{i}"].PutValue($"Item {i - 1}");
                cells[$"B{i}"].PutValue((i - 1) * 10);
            }

            // Group rows 2‑4 (zero‑based indices 1‑3) and hide the detail rows
            cells.GroupRows(1, 3, true);

            // Place the summary row above the grouped rows
            sheet.Outline.SummaryRowBelow = false;

            // Add a label shape that will act as a group heading.
            // Parameters: upper left row, upper left column, top offset (pixels), left offset (pixels), width, height
            Label groupLabel = sheet.Shapes.AddLabel(0, 0, 5, 5, 200, 30);
            groupLabel.Text = "Group 1 – Items 1‑3";
            groupLabel.Placement = PlacementType.FreeFloating; // Allows precise positioning
            groupLabel.IsHidden = false;
            groupLabel.IsLocked = false;

            // Adjust label position so it appears just above the grouped rows
            // UpperLeftRow = 0 (row 1), UpperLeftColumn = 0 (column A)
            groupLabel.UpperLeftRow = 0;
            groupLabel.UpperLeftColumn = 0;
            // Optional: fine‑tune pixel offsets
            groupLabel.Top = 2;   // pixels from the top of the cell
            groupLabel.Left = 2;  // pixels from the left of the cell

            // Save the workbook
            workbook.Save("GroupLabelDemo.xlsx");
        }
    }
}
