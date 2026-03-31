using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Add untyped (string) custom properties.
            // The Add(string, string) overload stores the value as PropertyType.String,
            // which is effectively an untyped textual metadata entry.
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            workbook.CustomDocumentProperties.Add("Version", "1.0.3");
            workbook.CustomDocumentProperties.Add("CreatedBy", "Jane Doe");
            workbook.CustomDocumentProperties.Add("Notes", "Initial release of the project.");

            // Optionally, verify that the properties were added.
            foreach (DocumentProperty prop in workbook.CustomDocumentProperties)
            {
                Console.WriteLine($"Name: {prop.Name}, Value: {prop.Value}, Type: {prop.Type}");
            }

            // Save the workbook to an XLSX file.
            workbook.Save("CustomPropertiesWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}