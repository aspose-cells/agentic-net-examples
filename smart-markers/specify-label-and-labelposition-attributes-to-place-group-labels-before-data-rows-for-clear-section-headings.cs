using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupLabelDemo
{
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
            cells["A2"].PutValue("Item A1");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("Item A2");
            cells["B3"].PutValue(150);
            cells["A4"].PutValue("Item A3");
            cells["B4"].PutValue(200);
            cells["A5"].PutValue("Item B1");
            cells["B5"].PutValue(120);
            cells["A6"].PutValue("Item B2");
            cells["B6"].PutValue(180);
            cells["A7"].PutValue("Item B3");
            cells["B7"].PutValue(220);

            // Group rows 2-4 (Item A) and rows 5-7 (Item B)
            cells.GroupRows(1, 3); // rows are zero‑based, so 1 = row 2
            cells.GroupRows(4, 6); // rows 5‑7

            // Place summary rows above the detail rows for each group
            sheet.Outline.SummaryRowBelow = false;

            // Add a label shape before the first group (above row 2)
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Label groupALabel = sheet.Shapes.AddLabel(0, 0, 5, 5, 150, 20);
            groupALabel.Text = "Group A Items";
            groupALabel.Placement = PlacementType.FreeFloating;
            groupALabel.Font.Size = 12;
            groupALabel.Font.IsBold = true;

            // Add a label shape before the second group (above row 5)
            Label groupBLabel = sheet.Shapes.AddLabel(3, 0, 5, 5, 150, 20);
            groupBLabel.Text = "Group B Items";
            groupBLabel.Placement = PlacementType.FreeFloating;
            groupBLabel.Font.Size = 12;
            groupBLabel.Font.IsBold = true;

            // Save the workbook
            workbook.Save("GroupLabelsDemo.xlsx");
        }
    }
}