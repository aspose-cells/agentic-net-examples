using System;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(12345);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Create OoxmlSaveOptions and configure compression (feature rule)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            // Use the highest compression level to minimize file size
            saveOptions.CompressionType = OoxmlCompressionType.Level9;
            // Disable exporting cell names to further reduce size
            saveOptions.ExportCellName = false;

            // Save the workbook to XLSX format with the specified options (lifecycle rule: save)
            workbook.Save("CompressedOutput.xlsx", saveOptions);

            Console.WriteLine("Workbook saved as 'CompressedOutput.xlsx' with maximum compression.");
        }
    }
}