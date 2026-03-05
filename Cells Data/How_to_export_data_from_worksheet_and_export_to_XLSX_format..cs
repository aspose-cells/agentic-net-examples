using System;
using System.Data;
using Aspose.Cells;

class ExportWorksheetToXlsx
{
    static void Main()
    {
        // Load an existing workbook (replace with your source file path)
        Workbook workbook = new Workbook("SourceWorkbook.xlsx");

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Export the range A1:D10 to a DataTable (including column names)
        // A1:D10 corresponds to rows 0-9 and columns 0-3 (zero‑based indexing)
        DataTable table = worksheet.Cells.ExportDataTable(0, 0, 10, 4, true);

        // Example: display the exported rows in the console
        Console.WriteLine("Exported DataTable rows:");
        foreach (DataRow row in table.Rows)
        {
            foreach (var item in row.ItemArray)
                Console.Write(item + "\t");
            Console.WriteLine();
        }

        // Save the workbook to XLSX format
        workbook.Save("ExportedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}