using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ContentTypePropertyMetadataDemo
    {
        public static void Run()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add new content type properties
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Add a simple string property (uses Add(string, string))
            int index1 = workbook.ContentTypeProperties.Add("ProjectName", "AsposeDemo");

            // Add a property with explicit type information (uses Add(string, string, string))
            int index2 = workbook.ContentTypeProperties.Add("CreatedOn", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"), "DateTime");

            // -------------------------------------------------
            // 2. Modify existing properties
            // -------------------------------------------------
            // Using indexer by index
            workbook.ContentTypeProperties[index1].IsNillable = true;          // allow empty value
            workbook.ContentTypeProperties[index1].Value = "Aspose.Cells";    // update value

            // Using indexer by name
            var createdOnProp = workbook.ContentTypeProperties["CreatedOn"];
            createdOnProp.IsNillable = false;                                 // disallow empty value
            createdOnProp.Value = DateTime.UtcNow.ToString("o");              // update to current UTC time

            // -------------------------------------------------
            // 3. Save the workbook with the new metadata
            // -------------------------------------------------
            workbook.Save("ContentTypePropertiesAdded.xlsx");

            // -------------------------------------------------
            // 4. Load an existing workbook and modify its content type properties
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook("ContentTypePropertiesAdded.xlsx");

            // Add another property to the loaded workbook
            loadedWorkbook.ContentTypeProperties.Add("Version", "1.2.3", "string");

            // Modify an existing property by name
            if (loadedWorkbook.ContentTypeProperties["ProjectName"] != null)
            {
                loadedWorkbook.ContentTypeProperties["ProjectName"].Value = "Aspose.Cells Updated";
            }

            // Save the changes to a new file
            loadedWorkbook.Save("ContentTypePropertiesModified.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ContentTypePropertyMetadataDemo.Run();
        }
    }
}