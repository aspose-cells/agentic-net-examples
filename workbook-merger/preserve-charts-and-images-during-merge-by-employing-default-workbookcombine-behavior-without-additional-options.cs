using System;
using Aspose.Cells;

class PreserveChartsAndImagesDuringMerge
{
    static void Main()
    {
        // Load the destination workbook (the one that will receive the other workbook's content)
        Workbook destWorkbook = new Workbook("Destination.xlsx");

        // Load the source workbook (the one whose content will be merged into the destination)
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Combine the source workbook into the destination workbook.
        // The default Combine method preserves charts, images, and other objects.
        destWorkbook.Combine(sourceWorkbook);

        // Save the merged workbook to a new file.
        destWorkbook.Save("CombinedResult.xlsx", SaveFormat.Xlsx);
    }
}