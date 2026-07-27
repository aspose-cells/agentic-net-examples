using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Source worksheet (first sheet)
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Populate source row (row index 0)
        srcSheet.Cells["A1"].PutValue("Item");
        srcSheet.Cells["B1"].PutValue(123);
        srcSheet.Cells["C1"].PutValue(DateTime.Now);

        // Destination worksheet (add a new sheet)
        Worksheet destSheet = workbook.Worksheets.Add("Destination");

        // Transfer the first row from source to destination
        destSheet.Cells.CopyRow(srcSheet.Cells, 0, 0);

        // Save the workbook
        workbook.Save("RowCopyExample.xlsx", SaveFormat.Xlsx);
    }
}

// Author: Aspose.Cells .NET example