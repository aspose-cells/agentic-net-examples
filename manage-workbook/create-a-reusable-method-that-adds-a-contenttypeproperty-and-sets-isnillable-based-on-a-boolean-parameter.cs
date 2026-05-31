using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

public static class ContentTypeHelper
{
    // Adds a content type property to the workbook and sets its IsNillable flag.
    public static void AddContentTypeProperty(Workbook workbook, string name, string value, string type, bool isNillable)
    {
        // Add the property; the overload with type returns the index of the new property.
        int index = workbook.ContentTypeProperties.Add(name, value, type);
        // Retrieve the property by index and set IsNillable.
        workbook.ContentTypeProperties[index].IsNillable = isNillable;
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook wb = new Workbook();

        // Add a property that can be empty.
        ContentTypeHelper.AddContentTypeProperty(wb, "Admin", "Aspose", "text", true);

        // Add a property that cannot be empty.
        ContentTypeHelper.AddContentTypeProperty(wb, "Version", "1.0", "number", false);

        // Save the workbook.
        wb.Save("ContentTypePropertiesDemo.xlsx");
    }
}