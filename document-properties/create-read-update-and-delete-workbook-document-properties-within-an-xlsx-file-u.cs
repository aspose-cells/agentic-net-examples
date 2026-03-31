using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class DocumentPropertiesDemo
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new workbook and set initial properties
        // -------------------------------------------------
        Workbook workbook = new Workbook();                     // create workbook
        // Built‑in properties
        workbook.BuiltInDocumentProperties["Author"].Value = "John Smith";
        workbook.BuiltInDocumentProperties["Title"].Value = "Initial Title";
        // Custom properties
        workbook.CustomDocumentProperties.Add("Project", "AsposeDemo");
        workbook.CustomDocumentProperties.Add("Revision", 1);
        workbook.CustomDocumentProperties.Add("Approved", false);
        // Save the workbook with the properties
        workbook.Save("PropertiesDemo.xlsx");

        // -------------------------------------------------
        // 2. Load the workbook and read the properties
        // -------------------------------------------------
        Workbook loaded = new Workbook("PropertiesDemo.xlsx"); // load workbook
        Console.WriteLine("=== Read Properties ===");
        Console.WriteLine("Author: " + loaded.BuiltInDocumentProperties["Author"].Value);
        Console.WriteLine("Title : " + loaded.BuiltInDocumentProperties["Title"].Value);
        Console.WriteLine("Project: " + loaded.CustomDocumentProperties["Project"].Value);
        Console.WriteLine("Revision: " + loaded.CustomDocumentProperties["Revision"].Value);
        Console.WriteLine("Approved: " + loaded.CustomDocumentProperties["Approved"].Value);

        // -------------------------------------------------
        // 3. Update existing properties
        // -------------------------------------------------
        loaded.BuiltInDocumentProperties["Title"].Value = "Updated Title";
        loaded.CustomDocumentProperties["Revision"].Value = 2;          // change number
        loaded.CustomDocumentProperties["Approved"].Value = true;      // change boolean

        // -------------------------------------------------
        // 4. Delete a custom property
        // -------------------------------------------------
        // The collection supports removal by name; if not, use RemoveAt(index)
        if (loaded.CustomDocumentProperties.Contains("Project"))
        {
            loaded.CustomDocumentProperties.Remove("Project");
        }

        // -------------------------------------------------
        // 5. Save the changes
        // -------------------------------------------------
        loaded.Save("PropertiesDemo_Updated.xlsx");

        // -------------------------------------------------
        // 6. Verify updates and deletion
        // -------------------------------------------------
        Workbook verify = new Workbook("PropertiesDemo_Updated.xlsx");
        Console.WriteLine("\n=== Verify After Update/Delete ===");
        Console.WriteLine("Title : " + verify.BuiltInDocumentProperties["Title"].Value);
        Console.WriteLine("Revision: " + verify.CustomDocumentProperties["Revision"].Value);
        Console.WriteLine("Approved: " + verify.CustomDocumentProperties["Approved"].Value);
        Console.WriteLine("Project exists? " + verify.CustomDocumentProperties.Contains("Project"));
    }
}