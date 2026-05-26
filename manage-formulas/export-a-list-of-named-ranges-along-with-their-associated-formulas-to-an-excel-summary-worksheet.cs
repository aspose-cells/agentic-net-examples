using System;
using Aspose.Cells;

class ExportNamedRanges
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and named ranges (optional)
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        dataSheet.Cells["A1"].PutValue(10);
        dataSheet.Cells["A2"].PutValue(20);
        dataSheet.Cells["B1"].PutValue(30);
        dataSheet.Cells["B2"].PutValue(40);
        dataSheet.Cells.CreateRange("A1:A2").Name = "FirstColumn";
        dataSheet.Cells.CreateRange("B1:B2").Name = "SecondColumn";

        // Add a summary worksheet
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[summaryIndex];
        summarySheet.Name = "NamedRangesSummary";

        // Write header
        summarySheet.Cells["A1"].PutValue("Name");
        summarySheet.Cells["B1"].PutValue("RefersTo");

        // Export named ranges and their formulas
        int row = 1; // zero‑based index, start after header
        foreach (Name name in workbook.Worksheets.Names)
        {
            summarySheet.Cells[row, 0].PutValue(name.Text);      // Named range name
            summarySheet.Cells[row, 1].PutValue(name.RefersTo); // Formula the name refers to
            row++;
        }

        // Save the workbook
        workbook.Save("NamedRangesSummary.xlsx");
    }
}