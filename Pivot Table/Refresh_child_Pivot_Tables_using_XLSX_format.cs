using System;
using System.IO;
using Aspose.Cells;

class RefreshChildPivotTables
{
    static void Main()
    {
        string inputPath = "ParentChildPivot.xlsx";

        // Create a sample workbook with a pivot table if the file does not exist
        if (!File.Exists(inputPath))
        {
            Workbook tempWb = new Workbook();
            Worksheet ws = tempWb.Worksheets[0];

            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(100);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(200);

            // Add a simple pivot table
            ws.PivotTables.Add("=A1:B3", "D1", "PivotTable1");

            tempWb.Save(inputPath);
        }

        // Load the workbook that contains parent and child pivot tables
        Workbook workbook = new Workbook(inputPath);

        // Example modification of source data (optional)
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["B2"].PutValue(1500); // change a source value

        // Refresh all pivot tables in the workbook, which also refreshes child pivot tables
        workbook.Worksheets.RefreshPivotTables();

        // Save the refreshed workbook
        workbook.Save("ParentChildPivot_Refreshed.xlsx");
    }
}