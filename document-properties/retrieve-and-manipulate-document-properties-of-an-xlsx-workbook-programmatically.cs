using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace DocumentPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook (creation rule)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 2. Manipulate built‑in document properties
            // -----------------------------------------------------------------
            // Set the Author property
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";

            // Set the Title property
            workbook.BuiltInDocumentProperties["Title"].Value = "Quarterly Report";

            // Set the Company property
            workbook.BuiltInDocumentProperties["Company"].Value = "Acme Corp";

            // -----------------------------------------------------------------
            // 3. Add custom document properties
            // -----------------------------------------------------------------
            // Add a string property
            workbook.CustomDocumentProperties.Add("Project", "Apollo");

            // Add a boolean property
            workbook.CustomDocumentProperties.Add("Reviewed", false);

            // Add an integer property
            workbook.CustomDocumentProperties.Add("Revision", 1);

            // -----------------------------------------------------------------
            // 4. Save the workbook to disk (save rule)
            // -----------------------------------------------------------------
            string filePath = "DocumentPropertiesDemo.xlsx";
            workbook.Save(filePath);

            // -----------------------------------------------------------------
            // 5. Load the workbook from the saved file (load rule)
            // -----------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);

            // -----------------------------------------------------------------
            // 6. Retrieve and display built‑in properties
            // -----------------------------------------------------------------
            Console.WriteLine("Built‑in Properties:");
            Console.WriteLine($"Author : {loadedWorkbook.BuiltInDocumentProperties["Author"].Value}");
            Console.WriteLine($"Title  : {loadedWorkbook.BuiltInDocumentProperties["Title"].Value}");
            Console.WriteLine($"Company: {loadedWorkbook.BuiltInDocumentProperties["Company"].Value}");

            // -----------------------------------------------------------------
            // 7. Retrieve and display custom properties
            // -----------------------------------------------------------------
            Console.WriteLine("\nCustom Properties:");
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name} = {prop.Value} ({prop.Type})");
            }

            // -----------------------------------------------------------------
            // 8. Modify a property and save again
            // -----------------------------------------------------------------
            // Mark the document as reviewed
            loadedWorkbook.CustomDocumentProperties["Reviewed"].Value = true;

            // Increment the revision number
            int rev = (int)loadedWorkbook.CustomDocumentProperties["Revision"].Value;
            loadedWorkbook.CustomDocumentProperties["Revision"].Value = rev + 1;

            // Save the updated workbook
            string updatedFilePath = "DocumentPropertiesDemo_Updated.xlsx";
            loadedWorkbook.Save(updatedFilePath);

            Console.WriteLine("\nProperties updated and workbook saved.");
        }
    }
}