using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerFormatting
{
    // Simple data class that will be bound to the smart markers
    public class Record
    {
        public DateTime Date { get; set; }
        public double Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook (lifecycle: create) ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Set up the template with formatted smart markers ----------
            // Header cells
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Amount");

            // Smart markers with formatting specifications
            // ?format=yyyy-MM-dd formats the DateTime value
            // ?format=Currency formats the numeric value as currency
            sheet.Cells["A2"].PutValue("&=$Date?format=yyyy-MM-dd");
            sheet.Cells["B2"].PutValue("&=$Amount?format=Currency");

            // ---------- Prepare the data source ----------
            List<Record> records = new List<Record>
            {
                new Record { Date = new DateTime(2023, 12, 25), Amount = 1999.99 },
                new Record { Date = DateTime.Today, Amount = 123.45 },
                new Record { Date = DateTime.Now.AddDays(7), Amount = 9876.54 }
            };

            // ---------- Process smart markers (lifecycle: load & process) ----------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the data source to a name that matches the smart marker table name ("Data")
            designer.SetDataSource("Data", records);
            // Process all smart markers in the workbook
            designer.Process();

            // ---------- Save the result (lifecycle: save) ----------
            workbook.Save("FormattedSmartMarkers.xlsx");
        }
    }
}