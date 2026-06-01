using System;
using Aspose.Cells;

namespace MergeWithTimestampDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load the first workbook (destination workbook)
            Workbook destWorkbook = new Workbook("Source1.xlsx");

            // Load the second workbook to be merged
            Workbook sourceWorkbook = new Workbook("Source2.xlsx");

            // Combine the second workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Add a new worksheet to hold the merge timestamp
            Worksheet timestampSheet = destWorkbook.Worksheets.Add("MergeTimestamp");

            // Write the exact date and time of the merge operation into cell A1
            timestampSheet.Cells["A1"].PutValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // Save the merged workbook with the timestamp worksheet
            destWorkbook.Save("MergedWithTimestamp.xlsx");
        }
    }
}