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
        // The third parameter specifies the property type; using "string" here
        workbook.ContentTypeProperties.Add("ProjectId", "MyProject123", "string");

        // Save the workbook to a file
        workbook.Save("ProjectWorkbook.xlsx");
    }
}