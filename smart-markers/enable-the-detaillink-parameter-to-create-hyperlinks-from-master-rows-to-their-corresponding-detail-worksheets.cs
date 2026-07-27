using System;
using Aspose.Cells;

class DetailLinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Master worksheet
        Worksheet masterSheet = workbook.Worksheets[0];
        masterSheet.Name = "Master";

        // Populate master data
        masterSheet.Cells["A1"].PutValue("ID");
        masterSheet.Cells["B1"].PutValue("Name");
        masterSheet.Cells["A2"].PutValue(1);
        masterSheet.Cells["B2"].PutValue("Item1");
        masterSheet.Cells["A3"].PutValue(2);
        masterSheet.Cells["B3"].PutValue("Item2");

        // Detail worksheet
        Worksheet detailSheet = workbook.Worksheets.Add("Detail");
        detailSheet.Cells["A1"].PutValue("Detail for Item1");
        detailSheet.Cells["A2"].PutValue("Detail for Item2");

        // Add hyperlinks from master rows to corresponding detail rows
        // Link from master row 2 (A2) to Detail!A1
        masterSheet.Hyperlinks.Add("A2", 1, 1, "Detail!A1");
        masterSheet.Hyperlinks[0].TextToDisplay = "View Detail 1";

        // Link from master row 3 (A3) to Detail!A2
        masterSheet.Hyperlinks.Add("A3", 1, 1, "Detail!A2");
        masterSheet.Hyperlinks[1].TextToDisplay = "View Detail 2";

        // Save the workbook
        workbook.Save("DetailLinkDemo.xlsx");
    }
}