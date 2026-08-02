using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class VerifyCoreProperties
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set initial built‑in (core) properties
        workbook.BuiltInDocumentProperties["Author"].Value = "Original Author";
        workbook.BuiltInDocumentProperties["Title"].Value = "Original Title";

        // Capture core properties before adding custom ones
        object originalAuthor = workbook.BuiltInDocumentProperties["Author"].Value;
        object originalTitle = workbook.BuiltInDocumentProperties["Title"].Value;

        // Add custom document properties
        workbook.CustomDocumentProperties.Add("Project", "AsposeDemo");
        workbook.CustomDocumentProperties.Add("Version", 1);
        workbook.CustomDocumentProperties.Add("Reviewed", true);
        workbook.CustomDocumentProperties.Add("CreatedOn", DateTime.Now);

        // Verify that core properties remain unchanged
        object afterAuthor = workbook.BuiltInDocumentProperties["Author"].Value;
        object afterTitle = workbook.BuiltInDocumentProperties["Title"].Value;

        Console.WriteLine("Author unchanged: " + originalAuthor.Equals(afterAuthor));
        Console.WriteLine("Title unchanged: " + originalTitle.Equals(afterTitle));

        // Save the workbook
        workbook.Save("VerifyCoreProperties.xlsx");
    }
}