using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access and modify built‑in document properties
        workbook.BuiltInDocumentProperties["Author"].Value = "John Smith";
        workbook.BuiltInDocumentProperties["Title"].Value = "Sample Workbook";

        // Add custom document properties
        workbook.CustomDocumentProperties.Add("Reviewed", true);
        workbook.CustomDocumentProperties.Add("Revision", 2);
        workbook.CustomDocumentProperties.Add("CreatedDate", DateTime.Now);

        // Display built‑in properties
        Console.WriteLine("Author: " + workbook.BuiltInDocumentProperties["Author"].Value);
        Console.WriteLine("Title: " + workbook.BuiltInDocumentProperties["Title"].Value);

        // Display custom properties
        foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
        {
            Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("DocumentPropertiesDemo.xlsx");
    }
}