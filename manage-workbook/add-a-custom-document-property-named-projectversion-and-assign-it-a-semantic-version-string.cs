using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddCustomDocumentProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProjectVersion" with a semantic version string
        workbook.CustomDocumentProperties.Add("ProjectVersion", "1.2.3");

        // Display the added property value (optional)
        Console.WriteLine("ProjectVersion: " + workbook.CustomDocumentProperties["ProjectVersion"].Value);

        // Save the workbook
        workbook.Save("ProjectVersionDemo.xlsx", SaveFormat.Xlsx);
    }
}