// Title: Check that built‑in workbook properties stay unchanged after adding custom document properties with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# keep original workbook author after adding custom properties | how to verify Excel built‑in metadata unchanged when using CustomDocumentProperties in .NET | C# Aspose.Cells compare built‑in document properties before and after adding custom properties | preserve workbook core properties while adding custom document properties with Aspose.Cells | Aspose.Cells verify that CreatedTime is not modified after adding custom properties
// Tags: Aspose.Cells verify built-in properties | add custom document properties preserve core metadata | C# compare workbook built-in properties | Excel core metadata unchanged Aspose.Cells | custom properties impact on built-in properties

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook or creates a new one, records its built‑in document properties (author, title, subject, keywords, comments, created time), adds several custom document properties, then checks that the original built‑in values are still the same, outputting whether the core properties remained unchanged.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if the template file exists; otherwise create a new workbook.
            Workbook workbook;
            string templatePath = "template.xlsx";

            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Capture the original core (built‑in) document properties.
            var builtInProps = workbook.BuiltInDocumentProperties;
            string originalAuthor = builtInProps.Author;
            string originalTitle = builtInProps.Title;
            string originalSubject = builtInProps.Subject;
            string originalKeywords = builtInProps.Keywords;
            string originalComments = builtInProps.Comments;
            DateTime originalCreatedTime = builtInProps.CreatedTime;

            // Add custom document properties.
            workbook.CustomDocumentProperties.Add("CustomProp1", "Value1");
            workbook.CustomDocumentProperties.Add("CustomProp2", 12345);
            workbook.CustomDocumentProperties.Add("CustomProp3", DateTime.Now);

            // Verify that core properties have not changed after adding custom properties.
            bool corePropertiesUnchanged =
                originalAuthor == builtInProps.Author &&
                originalTitle == builtInProps.Title &&
                originalSubject == builtInProps.Subject &&
                originalKeywords == builtInProps.Keywords &&
                originalComments == builtInProps.Comments &&
                originalCreatedTime == builtInProps.CreatedTime;

            Console.WriteLine("Core properties unchanged: " + corePropertiesUnchanged);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
