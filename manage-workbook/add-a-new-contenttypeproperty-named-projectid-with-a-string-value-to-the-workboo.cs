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
        // Using the Add(string name, string value) overload
        workbook.ContentTypeProperties.Add("ProjectId", "12345");

        // Save the workbook to a file
        workbook.Save("ProjectWorkbook.xlsx");
    }
}