using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data required for a PivotTable (date field is mandatory for a Timeline)
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Sales";
        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = 1200;
        cells["A3"].Value = new DateTime(2023, 2, 1);
        cells["B3"].Value = 1500;
        cells["A4"].Value = new DateTime(2023, 3, 1);
        cells["B4"].Value = 1800;

        // Create a PivotTable that will serve as the data source for the Timeline
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a TextBox shape that will contain tab‑separated text
        Shape textBox = sheet.Shapes.AddTextBox(5, 0, 5, 0, 300, 100);

        // Access the first paragraph of the TextBox
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];

        // Add tab stops: left‑aligned at 50 points, right‑aligned at 200 points
        paragraph.Stops.Add(TextTabAlignmentType.Left, 50.0);
        paragraph.Stops.Add(TextTabAlignmentType.Right, 200.0);

        // Set the text using TAB characters to align with the defined stops
        textBox.TextBody.Text = "Product\tRevenue";

        // Save the workbook
        workbook.Save("TimelineWithTabFormat.xlsx");
    }
}