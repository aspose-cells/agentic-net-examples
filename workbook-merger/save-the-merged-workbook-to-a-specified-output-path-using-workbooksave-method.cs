using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the source workbook and add sample data
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

        // Create the destination workbook (XLSX format) and add sample data
        Workbook destinationWorkbook = new Workbook(FileFormatType.Xlsx);
        destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

        // Merge the source workbook into the destination workbook
        destinationWorkbook.Combine(sourceWorkbook);

        // Specify the output file path
        string outputPath = "CombinedWorkbook.xlsx";

        // Save the merged workbook using the Save(string, SaveFormat) method
        destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);
    }
}