using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define an array of tags to categorize the workbook content
        string[] tags = new string[] { "Finance", "Quarterly", "Confidential" };

        // Convert the string array to a single string (comma‑separated) for storage
        string tagsValue = string.Join(",", tags);

        // Add a custom document property named "Tags" with the serialized array value
        workbook.CustomDocumentProperties.Add("Tags", tagsValue);

        // Save the workbook to a file
        workbook.Save("WorkbookWithTags.xlsx");
    }
}