using System;
using System.Text;
using Aspose.Cells;

class ExportCsvWithBom
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(25);

        // Set up CSV save options with UTF-8 encoding (includes BOM)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Encoding = Encoding.UTF8; // ensures a Byte Order Mark is written

        // Save the workbook as a CSV file with the specified options
        workbook.Save("output_with_bom.csv", saveOptions);
    }
}