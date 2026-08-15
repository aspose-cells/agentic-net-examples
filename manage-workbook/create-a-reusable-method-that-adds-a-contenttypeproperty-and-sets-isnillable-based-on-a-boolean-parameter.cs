// Title: Reusable C# method to add a ContentTypeProperty and set IsNillable in Aspose.Cells
// Description: Demonstrates a static helper that adds a custom ContentTypeProperty to an Aspose.Cells Workbook, assigns its IsNillable flag based on a Boolean argument, and saves the workbook. Ideal for centralising metadata handling in .NET Excel generation.
// Keywords: Aspose.Cells ContentTypeProperty | IsNillable flag C# | Workbook metadata helper | Add custom property Aspose.Cells | Reusable Excel utility .NET | ContentTypeProperties.Add example | C# Aspose.Cells tutorial
// Common Searches: how to set IsNillable on ContentTypeProperty Aspose.Cells | C# method to add custom content type property to workbook | Aspose.Cells reusable helper for workbook metadata | add text property with nillable flag in Excel using Aspose | sample code for ContentTypeProperties.Add in .NET
// Developer Intent: Create a single method that adds a ContentTypeProperty to a workbook and configures its IsNillable attribute according to a supplied boolean.
// Use Cases: Standardise optional metadata fields (e.g., author, department) across generated reports. | Record version or identifier properties that must never be null in compliance documents. | Provide a utility callable from multiple services to attach custom text or numeric properties with configurable nillability.
// AI Prompts: Generate a C# extension method for Aspose.Cells that adds a ContentTypeProperty and sets IsNillable based on a bool parameter. | Write unit tests using xUnit to verify that the IsNillable flag is true when the helper receives true and false otherwise. | Extend the helper to accept an optional namespace URI and default value for the new property.

using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates a static helper that adds a custom ContentTypeProperty to an Aspose.Cells Workbook, assigns its IsNillable flag based on a Boolean argument, and saves the workbook. Ideal for centralising metadata handling in .NET Excel generation.
public static class ContentTypeHelper
{
    // Adds a content type property to the workbook and sets its IsNillable flag.
    public static void AddContentTypeProperty(Workbook workbook, string name, string value, string type, bool isNillable)
    {
        // Add the property; the method returns the index of the newly added property.
        int index = workbook.ContentTypeProperties.Add(name, value, type);
        // Retrieve the property via the indexer and set the IsNillable property.
        workbook.ContentTypeProperties[index].IsNillable = isNillable;
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook instance.
        Workbook wb = new Workbook();

        // Add a content type property with IsNillable set to true.
        ContentTypeHelper.AddContentTypeProperty(wb, "Admin", "Aspose", "text", true);

        // Add another content type property with IsNillable set to false.
        ContentTypeHelper.AddContentTypeProperty(wb, "Version", "1.0", "number", false);

        // Save the workbook to a file.
        wb.Save("ContentTypePropertiesDemo.xlsx");
    }
}
