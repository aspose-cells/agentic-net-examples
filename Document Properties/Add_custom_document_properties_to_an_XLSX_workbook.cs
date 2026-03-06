using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add custom document properties of various data types
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");          // string
            workbook.CustomDocumentProperties.Add("Revision", 3);                  // int
            workbook.CustomDocumentProperties.Add("CreatedOn", DateTime.Now);     // DateTime
            workbook.CustomDocumentProperties.Add("IsApproved", true);           // bool
            workbook.CustomDocumentProperties.Add("Score", 4.75);                 // double

            // Display the added properties (optional)
            foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }

            // Save the workbook with the custom properties
            workbook.Save("CustomPropertiesDemo.xlsx");
        }
    }
}