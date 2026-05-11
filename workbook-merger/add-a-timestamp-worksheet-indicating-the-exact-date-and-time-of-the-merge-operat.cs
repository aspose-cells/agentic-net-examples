using System;
using Aspose.Cells;

class MergeWithTimestamp
{
    static void Main()
    {
        // Load the destination workbook (the one that will receive the merge)
        Workbook destWorkbook = new Workbook("Dest.xlsx");

        // Load the source workbook (the one to be merged into the destination)
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Merge the source workbook into the destination workbook
        destWorkbook.Combine(sourceWorkbook);

        // Add a new worksheet named "Timestamp" to record the merge time
        Worksheet timestampSheet = destWorkbook.Worksheets.Add("Timestamp");

        // Write a label
        timestampSheet.Cells["A1"].PutValue("Merge performed at:");

        // Write the exact date and time of the merge operation
        timestampSheet.Cells["B1"].PutValue(DateTime.Now);

        // Save the merged workbook with the timestamp worksheet
        destWorkbook.Save("MergedWithTimestamp.xlsx", SaveFormat.Xlsx);
    }
}