using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add content type properties with different IsNillable settings
        int index1 = workbook.ContentTypeProperties.Add("Property1", "Sample Value");
        workbook.ContentTypeProperties[index1].IsNillable = true;

        int index2 = workbook.ContentTypeProperties.Add("Property2", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"), "DateTime");
        workbook.ContentTypeProperties[index2].IsNillable = false;

        int index3 = workbook.ContentTypeProperties.Add("Property3", "Another Value", "text");
        workbook.ContentTypeProperties[index3].IsNillable = true;

        // Generate a summary indicating which properties are nillable
        Console.WriteLine("ContentTypeProperty Nillable Summary:");
        foreach (ContentTypeProperty prop in workbook.ContentTypeProperties)
        {
            Console.WriteLine($"- {prop.Name}: {(prop.IsNillable ? "Nillable" : "Not Nillable")}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ContentTypePropertiesSummary.xlsx");
    }
}