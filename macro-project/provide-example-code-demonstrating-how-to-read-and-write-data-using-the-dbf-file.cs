using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Loading;

class DbfReadWriteDemo
{
    static void Main()
    {
        // -------------------- Create a workbook and add sample data --------------------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");

        // Data rows
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");

        // -------------------- Save the workbook as a DBF file --------------------
        DbfSaveOptions saveOptions = new DbfSaveOptions();
        // Export all values as strings (optional, demonstrates the property)
        saveOptions.ExportAsString = true;
        workbook.Save("sample.dbf", saveOptions);

        // -------------------- Load the DBF file back into a workbook --------------------
        DbfLoadOptions loadOptions = new DbfLoadOptions();
        Workbook dbfWorkbook = new Workbook("sample.dbf", loadOptions);
        Worksheet dbfWorksheet = dbfWorkbook.Worksheets[0];

        // -------------------- Read and display the loaded data --------------------
        Console.WriteLine("Data loaded from DBF file:");
        Console.WriteLine($"{dbfWorksheet.Cells["A1"].StringValue}\t{dbfWorksheet.Cells["B1"].StringValue}");
        Console.WriteLine($"{dbfWorksheet.Cells["A2"].StringValue}\t{dbfWorksheet.Cells["B2"].StringValue}");
        Console.WriteLine($"{dbfWorksheet.Cells["A3"].StringValue}\t{dbfWorksheet.Cells["B3"].StringValue}");

        // -------------------- Modify a cell value --------------------
        dbfWorksheet.Cells["B2"].PutValue("Alice Smith");

        // -------------------- Save the modified workbook to another format (e.g., XLSX) --------------------
        dbfWorkbook.Save("modified.xlsx", SaveFormat.Xlsx);
    }
}