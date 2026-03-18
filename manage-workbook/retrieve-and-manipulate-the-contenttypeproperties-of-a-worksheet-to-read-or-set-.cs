using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class ContentTypePropertiesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a string content type property
        int idxProject = wb.ContentTypeProperties.Add("Project", "AsposeDemo", "string");
        wb.ContentTypeProperties[idxProject].IsNillable = true;

        // Add a DateTime content type property with explicit type
        int idxCreated = wb.ContentTypeProperties.Add(
            "CreatedOn",
            DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
            "DateTime");
        wb.ContentTypeProperties[idxCreated].IsNillable = false;

        // Retrieve property by index
        ContentTypeProperty propByIndex = wb.ContentTypeProperties[0];
        Console.WriteLine($"Index 0 -> Name: {propByIndex.Name}, Value: {propByIndex.Value}, Type: {propByIndex.Type}, IsNillable: {propByIndex.IsNillable}");

        // Retrieve property by name
        ContentTypeProperty propByName = wb.ContentTypeProperties["CreatedOn"];
        Console.WriteLine($"Name 'CreatedOn' -> Value: {propByName.Value}");

        // Modify the DateTime property's value and nillable flag
        propByName.Value = DateTime.UtcNow.ToString("o");
        propByName.IsNillable = true;

        // Save the workbook to a file
        string filePath = "ContentTypePropertiesDemo.xlsx";
        wb.Save(filePath);

        // Load the workbook from the file and read back a property
        Workbook loadedWb = new Workbook(filePath);
        ContentTypeProperty loadedProp = loadedWb.ContentTypeProperties["Project"];
        Console.WriteLine($"Loaded Property -> Name: {loadedProp.Name}, Value: {loadedProp.Value}");
    }
}