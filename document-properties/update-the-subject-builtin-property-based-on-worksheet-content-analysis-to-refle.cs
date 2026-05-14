using System;
using Aspose.Cells;

class UpdateSubjectBasedOnContent
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data – the first non‑empty cell in the first row will be treated as the main topic
        sheet.Cells["A1"].PutValue("Sales Report Q1 2024");
        sheet.Cells["A2"].PutValue("Region");
        sheet.Cells["B2"].PutValue("Revenue");
        sheet.Cells["A3"].PutValue("North");
        sheet.Cells["B3"].PutValue(120000);
        sheet.Cells["A4"].PutValue("South");
        sheet.Cells["B4"].PutValue(95000);

        // Analyze the first row to find the first non‑empty cell (main topic)
        string mainTopic = string.Empty;
        int lastColumn = sheet.Cells.MaxColumn;
        for (int col = 0; col <= lastColumn; col++)
        {
            var cell = sheet.Cells[0, col];
            if (cell != null && cell.Value != null && !string.IsNullOrWhiteSpace(cell.StringValue))
            {
                mainTopic = cell.StringValue;
                break;
            }
        }

        // Update the Subject built‑in document property if a topic was identified
        if (!string.IsNullOrEmpty(mainTopic))
        {
            workbook.BuiltInDocumentProperties.Subject = mainTopic;
        }

        // Save the workbook
        workbook.Save("UpdatedSubject.xlsx", SaveFormat.Xlsx);
    }
}