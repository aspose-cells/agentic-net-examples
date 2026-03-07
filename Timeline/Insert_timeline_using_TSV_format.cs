using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineTSV
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare TSV data (tab‑separated values)
            string tsvData =
                "Date\tValue\n" +
                "2023-01-01\t100\n" +
                "2023-01-15\t150\n" +
                "2023-02-01\t200";

            // Convert TSV string to a memory stream
            using (MemoryStream tsvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(tsvData)))
            {
                // Import TSV data into the worksheet starting at cell A1 (row 0, column 0)
                // Use tab character as the separator, enable numeric conversion
                cells.ImportCSV(tsvStream, "\t", true, 0, 0);
            }

            // Create a pivot table based on the imported data (range A1:B4)
            // Place the pivot table starting at cell D1
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Date as row field, Value as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table, positioned at cell F1, based on the "Date" field
            sheet.Timelines.Add(pivot, "F1", "Date");

            // Save the workbook
            workbook.Save("TimelineFromTSV.xlsx");
        }
    }
}