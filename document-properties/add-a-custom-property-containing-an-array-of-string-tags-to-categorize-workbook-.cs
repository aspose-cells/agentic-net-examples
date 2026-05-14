using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomTagProperty
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define an array of tags to categorize the workbook content
            string[] tags = new[] { "Finance", "Quarterly", "Confidential" };

            // Convert the array to a single string (comma‑separated) and add it as a custom document property
            // CustomDocumentProperties.Add(string name, string value) stores a string property
            workbook.CustomDocumentProperties.Add("Tags", string.Join(",", tags));

            // Save the workbook to a file (the property will be persisted)
            workbook.Save("WorkbookWithTags.xlsx");
        }
    }
}