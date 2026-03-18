using System;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Compression Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure OOXML save options with a custom compression level (feature: compression)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CompressionType = OoxmlCompressionType.Level9; // Highest compression

            // Save the workbook to XLSX using the configured options (lifecycle: save)
            workbook.Save("CompressedOutput.xlsx", saveOptions);
        }
    }
}