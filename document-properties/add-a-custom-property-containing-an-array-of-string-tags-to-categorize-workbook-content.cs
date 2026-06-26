using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddTagsCustomProperty
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define the tags to categorize the workbook content
        string[] tags = new string[] { "Finance", "2023", "Q1" };

        // Convert the string array to a single delimited string (comma‑separated)
        string tagsValue = string.Join(",", tags);

        // Add a custom document property named "Tags" with the delimited string value
        // This uses the Add(string, string) overload which creates a PropertyType.String property
        workbook.CustomDocumentProperties.Add("Tags", tagsValue);

        // Save the workbook to a file
        workbook.Save("TaggedWorkbook.xlsx");
    }
}