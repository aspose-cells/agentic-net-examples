using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace DocumentPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Built‑in document properties (read/write)
            // -------------------------------------------------
            // Set Author and Title built‑in properties
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Sales Report Q1";

            // -------------------------------------------------
            // Custom document properties (add new entries)
            // -------------------------------------------------
            // Add a string property
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            // Add an integer property
            workbook.CustomDocumentProperties.Add("Revision", 3);
            // Add a DateTime property
            workbook.CustomDocumentProperties.Add("GeneratedOn", DateTime.Now);
            // Add a Boolean property
            workbook.CustomDocumentProperties.Add("Approved", true);

            // -------------------------------------------------
            // Save the workbook (lifecycle save)
            // -------------------------------------------------
            string filePath = "DocumentPropertiesDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Load the workbook back to verify properties
            // -------------------------------------------------
            Workbook loaded = new Workbook(filePath);

            // Display built‑in properties
            Console.WriteLine("Author: " + loaded.BuiltInDocumentProperties["Author"].Value);
            Console.WriteLine("Title : " + loaded.BuiltInDocumentProperties["Title"].Value);

            // Display custom properties
            foreach (DocumentProperty prop in loaded.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }
        }
    }
}