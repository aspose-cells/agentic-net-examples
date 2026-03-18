using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsContentTypeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Add content type properties to the workbook
            // -------------------------------------------------
            // Simple string property
            int idx1 = workbook.ContentTypeProperties.Add("Author", "John Doe", "string");
            // DateTime property (ISO 8601 format)
            int idx2 = workbook.ContentTypeProperties.Add("Created", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"), "DateTime");
            // Boolean property stored as string
            int idx3 = workbook.ContentTypeProperties.Add("IsApproved", "true", "boolean");

            // -------------------------------------------------
            // Modify properties after adding
            // -------------------------------------------------
            // Set IsNillable flag
            workbook.ContentTypeProperties[idx1].IsNillable = false;
            workbook.ContentTypeProperties[idx2].IsNillable = true;
            workbook.ContentTypeProperties[idx3].IsNillable = false;

            // Change the value of an existing property using the name indexer
            ContentTypeProperty authorProp = workbook.ContentTypeProperties["Author"];
            authorProp.Value = "Jane Smith";

            // Change the type of a property (e.g., make IsApproved a string)
            ContentTypeProperty approvedProp = workbook.ContentTypeProperties["IsApproved"];
            approvedProp.Type = "string";

            // -------------------------------------------------
            // Retrieve and display all content type properties
            // -------------------------------------------------
            Console.WriteLine("Content Type Properties in the workbook:");
            foreach (ContentTypeProperty prop in workbook.ContentTypeProperties)
            {
                Console.WriteLine($"Name: {prop.Name}, Value: {prop.Value}, Type: {prop.Type}, IsNillable: {prop.IsNillable}");
            }

            // -------------------------------------------------
            // Save the workbook to a file
            // -------------------------------------------------
            string filePath = "ContentTypeDemo.xlsx";
            workbook.Save(filePath);

            // -------------------------------------------------
            // Load the workbook back and verify the properties
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);
            Console.WriteLine("\nProperties after reloading the workbook:");
            foreach (ContentTypeProperty prop in loadedWorkbook.ContentTypeProperties)
            {
                Console.WriteLine($"Name: {prop.Name}, Value: {prop.Value}, Type: {prop.Type}, IsNillable: {prop.IsNillable}");
            }
        }
    }
}