using System;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data (optional, just to have some content)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Compression Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create OOXML save options and set the highest compression level (Level9)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CompressionType = OoxmlCompressionType.Level9;

            // Save the workbook as XLSX using the configured compression options
            workbook.Save("CompressedWorkbook.xlsx", saveOptions);
        }
    }
}