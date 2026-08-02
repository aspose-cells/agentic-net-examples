using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

public static class ContentTypeHelper
{
    // Adds a content type property to the workbook and sets its IsNillable flag.
    public static void AddContentTypeProperty(Workbook workbook, string name, string value, string type, bool isNillable)
    {
        // Add the property (returns the index of the newly added property)
        int index = workbook.ContentTypeProperties.Add(name, value, type);
        // Retrieve the property by index and set the IsNillable property
        workbook.ContentTypeProperties[index].IsNillable = isNillable;
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook wb = new Workbook();

        // Add a content type property with IsNillable set to true
        ContentTypeHelper.AddContentTypeProperty(wb, "Admin", "Aspose", "text", true);

        // Add another content type property with IsNillable set to false
        ContentTypeHelper.AddContentTypeProperty(wb, "Version", "1.0", "number", false);

        // Save the workbook to a file
        wb.Save("ContentTypePropertiesDemo.xlsx");
    }
}