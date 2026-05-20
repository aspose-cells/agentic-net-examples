using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ContentTypePropertyNillableSummary
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample content type properties with different IsNillable settings
        int idx1 = workbook.ContentTypeProperties.Add("PropA", "ValueA", "text");
        workbook.ContentTypeProperties[idx1].IsNillable = true;

        int idx2 = workbook.ContentTypeProperties.Add("PropB", "123", "number");
        workbook.ContentTypeProperties[idx2].IsNillable = false;

        int idx3 = workbook.ContentTypeProperties.Add("PropC", "", "string");
        workbook.ContentTypeProperties[idx3].IsNillable = true;

        // Generate a summary indicating which properties are nillable
        Console.WriteLine("ContentTypeProperty Nillable Summary:");
        foreach (ContentTypeProperty prop in workbook.ContentTypeProperties)
        {
            string status = prop.IsNillable ? "Nillable" : "Not Nillable";
            Console.WriteLine($"- {prop.Name}: {status}");
        }

        // Save the workbook (optional)
        workbook.Save("ContentTypePropertiesSummary.xlsx");
    }
}