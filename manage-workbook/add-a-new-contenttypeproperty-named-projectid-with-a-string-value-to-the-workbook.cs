using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a ContentTypeProperty named "ProjectId" with a string value
        // The third parameter specifies the property type ("string")
        workbook.ContentTypeProperties.Add("ProjectId", "PRJ-001", "string");

        // Save the workbook
        workbook.Save("ProjectWorkbook.xlsx");
    }
}