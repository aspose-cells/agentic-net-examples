using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the source workbook and add some data.
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

        // Create the destination workbook (XLSX format) and add its own data.
        Workbook destinationWorkbook = new Workbook(FileFormatType.Xlsx);
        destinationWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

        // Combine the source workbook into the destination workbook.
        destinationWorkbook.Combine(sourceWorkbook);

        // Persist the combined workbook to an XLSX file using the provided Save API.
        destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}