using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(12345);
        sheet.Cells["A3"].PutValue(DateTime.Now);

        // Configure save options to use the highest OOXML compression level (Level9)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level9;

        // Save the workbook with the specified compression options
        workbook.Save("CompressedWorkbook.xlsx", saveOptions);
    }
}