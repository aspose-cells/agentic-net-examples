using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- Create a workbook (template) ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");

        // Smart markers for data rows
        sheet.Cells["A2"].PutValue("&=$Data.Name");
        sheet.Cells["B2"].PutValue("&=$Data.Score");

        // Mark the range that contains smart markers
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // ---------- Prepare data source with null values ----------
        var data = new List<Person>
        {
            new Person { Name = "Alice",   Score = 85 },
            new Person { Name = null,     Score = 90 },   // Name is null
            new Person { Name = "Charlie",Score = null } // Score is null
        };

        // ---------- Process smart markers ----------
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            // Optional: treat empty strings as null
            UpdateEmptyStringAsNull = true
        };
        designer.SetDataSource("Data", data);
        designer.Process();

        // ---------- Apply AutoFilter to exclude rows with null/blank Name ----------
        // The range includes header and all possible data rows (max 4 rows here)
        sheet.AutoFilter.Range = "A1:B4";
        // Show only rows where column A (Name) is not blank
        sheet.AutoFilter.MatchNonBlanks(0);
        sheet.AutoFilter.Refresh();

        // ---------- Save the workbook ----------
        workbook.Save("SmartMarkerFilteredOutput.xlsx");
    }

    // Simple POCO used as data source
    public class Person
    {
        public string Name { get; set; }
        public int? Score { get; set; }
    }
}