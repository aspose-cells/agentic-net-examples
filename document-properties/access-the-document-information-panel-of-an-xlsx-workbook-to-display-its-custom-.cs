using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (uses the create rule)
            Workbook workbook = new Workbook();

            // Add some custom document properties
            workbook.CustomDocumentProperties.Add("Project", "Aspose.Cells Demo");
            workbook.CustomDocumentProperties.Add("Version", 1);
            workbook.CustomDocumentProperties.Add("CreatedOn", DateTime.Now);
            workbook.CustomDocumentProperties.Add("IsApproved", true);

            // Save the workbook to a file (uses the save rule)
            string filePath = "CustomPropertiesDemo.xlsx";
            workbook.Save(filePath);

            // Load the workbook from the file (uses the load rule)
            Workbook loadedWorkbook = new Workbook(filePath);

            // Access and display all custom document properties
            Console.WriteLine("Custom Document Properties:");
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                // Display property name, value and type
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }
        }
    }
}