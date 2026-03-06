using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomDocPropsDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add custom document properties (string, int, DateTime, bool, double)
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("CreatedOn", DateTime.Now);
            workbook.CustomDocumentProperties.Add("IsApproved", false);
            workbook.CustomDocumentProperties.Add("Score", 8.75);

            // Remove a property by name
            workbook.CustomDocumentProperties.Remove("IsApproved");

            // Save the workbook (lifecycle: save)
            string outputPath = "CustomPropertiesDemo.xlsx";
            workbook.Save(outputPath);

            // Load the saved workbook (lifecycle: load)
            Workbook loadedWorkbook = new Workbook(outputPath);

            // Verify remaining properties
            Console.WriteLine("Custom Document Properties after removal:");
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }

            // Add another property to the loaded workbook
            loadedWorkbook.CustomDocumentProperties.Add("ReviewedBy", "Jane Doe");

            // Save the updated workbook
            string updatedPath = "CustomPropertiesDemo_Updated.xlsx";
            loadedWorkbook.Save(updatedPath);
        }
    }
}