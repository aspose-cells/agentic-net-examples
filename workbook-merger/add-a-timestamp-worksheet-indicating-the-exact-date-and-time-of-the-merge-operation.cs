using System;
using Aspose.Cells;

class MergeWithTimestamp
{
    static void Main()
    {
        // Load the primary workbook (destination)
        Workbook destWorkbook = new Workbook("Source1.xlsx");

        // Load the workbook to be merged into the destination
        Workbook sourceWorkbook = new Workbook("Source2.xlsx");

        // Merge the source workbook into the destination workbook
        destWorkbook.Combine(sourceWorkbook);

        // Add a new worksheet to hold the merge timestamp
        Worksheet timestampSheet = destWorkbook.Worksheets.Add("MergeTimestamp");

        // Write a label and the exact date/time of the merge operation
        timestampSheet.Cells["A1"].PutValue("Merge performed at:");
        timestampSheet.Cells["B1"].PutValue(DateTime.Now);

        // Save the combined workbook with the timestamp worksheet
        destWorkbook.Save("MergedWithTimestamp.xlsx", SaveFormat.Xlsx);
    }
}