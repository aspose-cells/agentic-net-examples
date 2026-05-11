using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the destination workbook (the one that will receive the other workbook's content)
        Workbook destinationWorkbook = new Workbook("Destination.xlsx");

        // Load the source workbook (the one whose charts, images, etc. will be merged)
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Combine the source workbook into the destination workbook.
        // The default Combine method preserves charts, images, and other objects.
        destinationWorkbook.Combine(sourceWorkbook);

        // Save the merged workbook.
        destinationWorkbook.Save("MergedResult.xlsx", SaveFormat.Xlsx);
    }
}