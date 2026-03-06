using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace DocumentPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory Excel file)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Built‑in document properties (read‑write)
            // -------------------------------------------------
            // Set Author and Title built‑in properties
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Sales Report Q1";

            // -------------------------------------------------
            // Custom document properties (user‑defined)
            // -------------------------------------------------
            // Add a string property
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            // Add an integer property
            workbook.CustomDocumentProperties.Add("Revision", 3);
            // Add a DateTime property
            workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);
            // Add a Boolean property
            workbook.CustomDocumentProperties.Add("Approved", true);
            // Add a double (float) property
            workbook.CustomDocumentProperties.Add("Score", 87.5);

            // -------------------------------------------------
            // Save the workbook to an XLSX file
            // -------------------------------------------------
            string outputPath = "DocumentPropertiesDemo.xlsx";
            workbook.Save(outputPath);

            // -------------------------------------------------
            // Load the saved workbook to verify properties
            // -------------------------------------------------
            Workbook loaded = new Workbook(outputPath);

            // Display built‑in properties
            Console.WriteLine("Built‑in Properties:");
            Console.WriteLine($"Author: {loaded.BuiltInDocumentProperties["Author"].Value}");
            Console.WriteLine($"Title : {loaded.BuiltInDocumentProperties["Title"].Value}");

            // Display custom properties
            Console.WriteLine("\nCustom Properties:");
            foreach (DocumentProperty prop in loaded.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
            }
        }
    }
}