using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddCustomPropertiesWithoutType
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add custom properties without specifying a type (uses Worksheet.CustomProperties)
        sheet.CustomProperties.Add("ProjectName", "AsposeDemo");
        sheet.CustomProperties.Add("Version", "1.0");
        sheet.CustomProperties.Add("CreatedOn", DateTime.Now.ToString());

        // Save the workbook as XLSX
        workbook.Save("CustomPropertiesWithoutType.xlsx", SaveFormat.Xlsx);
    }
}