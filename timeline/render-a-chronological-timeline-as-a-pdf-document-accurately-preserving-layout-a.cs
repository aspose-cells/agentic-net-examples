using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

namespace TimelinePdfDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate sample data (Date, Category, Value)
            // ------------------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Value");

            // Apply date format to the Date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Short date format

            DateTime start = new DateTime(2023, 1, 1);
            string[] categories = { "Food", "Transport", "Entertainment" };
            Random rnd = new Random();

            for (int i = 0; i < 30; i++)
            {
                // Date column
                Cell dateCell = cells[i + 2, 0];
                dateCell.PutValue(start.AddDays(i));
                dateCell.SetStyle(dateStyle);

                // Category column
                cells[i + 2, 1].PutValue(categories[i % categories.Length]);

                // Value column (random)
                cells[i + 2, 2].PutValue(rnd.Next(10, 100));
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable based on the data
            // ------------------------------------------------------------
            // Define the source range (A1:C31) and the destination cell (E3)
            int pivotIndex = sheet.PivotTables.Add("A1:C31", "E3", "TimelinePivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields: Date to Row area, Category to Column area, Value to Data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Column, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline control linked to the PivotTable's Date field
            // ------------------------------------------------------------
            // Place the Timeline with its upper‑left corner at cell G1
            sheet.Timelines.Add(pivot, "G1", "Date");

            // Optional: customize the Timeline appearance
            Timeline timeline = sheet.Timelines[0];
            timeline.Caption = "Spending Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;

            // ------------------------------------------------------------
            // 4. Save the workbook as a PDF, preserving layout and structure
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            workbook.Save("ChronologicalTimeline.pdf", pdfOptions);
        }
    }
}