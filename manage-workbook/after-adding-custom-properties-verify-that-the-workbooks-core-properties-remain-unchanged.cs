using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided Workbook constructor rule)
            Workbook workbook = new Workbook();

            // Set some built‑in document properties that we will later verify remain unchanged
            workbook.BuiltInDocumentProperties["Author"].Value = "Jane Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Original Title";

            // Capture the original built‑in property values
            object originalAuthor = workbook.BuiltInDocumentProperties["Author"].Value;
            object originalTitle = workbook.BuiltInDocumentProperties["Title"].Value;

            // Add custom document properties (uses the provided CustomDocumentProperties.Add rule)
            workbook.CustomDocumentProperties.Add("Project", "AsposeDemo");
            workbook.CustomDocumentProperties.Add("Version", 1);
            workbook.CustomDocumentProperties.Add("Reviewed", true);
            workbook.CustomDocumentProperties.Add("GeneratedOn", DateTime.Now);

            // Verify that built‑in properties have not changed after adding custom properties
            bool authorUnchanged = Equals(originalAuthor, workbook.BuiltInDocumentProperties["Author"].Value);
            bool titleUnchanged = Equals(originalTitle, workbook.BuiltInDocumentProperties["Title"].Value);

            Console.WriteLine($"Author unchanged: {authorUnchanged}");
            Console.WriteLine($"Title unchanged: {titleUnchanged}");

            // Save the workbook (uses the provided Save method rule)
            workbook.Save("CustomPropertiesVerification.xlsx");

            // Optional: Load the saved workbook to double‑check the properties
            Workbook loaded = new Workbook("CustomPropertiesVerification.xlsx");
            Console.WriteLine($"Loaded Author: {loaded.BuiltInDocumentProperties["Author"].Value}");
            Console.WriteLine($"Loaded Title: {loaded.BuiltInDocumentProperties["Title"].Value}");
            Console.WriteLine($"Loaded Custom Property 'Project': {loaded.CustomDocumentProperties["Project"].Value}");
        }
    }
}