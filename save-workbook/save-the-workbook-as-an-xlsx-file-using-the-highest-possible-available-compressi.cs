using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Compression Test");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure OOXML save options with the highest compression level
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level9; // Best compression (slowest)

        // Save the workbook as XLSX using the specified options
        workbook.Save("CompressedWorkbook.xlsx", saveOptions);
    }
}