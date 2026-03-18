using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add user‑defined (custom) document properties
            // These properties will be visible in the Document Information Panel of Excel
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            workbook.CustomDocumentProperties.Add("Reviewed", true);
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("LastReviewed", DateTime.Now);

            // Optionally display the added properties in the console
            Console.WriteLine("Custom Document Properties:");
            foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomPropertiesDemo.xlsx");
        }
    }
}