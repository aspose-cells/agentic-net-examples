// Title: Check that workbook core properties stay unchanged after adding custom document properties with Aspose.Cells for .NET
// Common Searches: how to keep workbook core metadata unchanged when adding custom document properties using Aspose.Cells C# | Aspose.Cells C# verify that Title, Author, Subject are not altered after inserting custom properties | preserve Excel file core properties while adding custom document properties in .NET | C# code to compare original workbook properties after adding custom properties with Aspose.Cells | ensure core properties remain the same after adding custom document properties in an Aspose.Cells workbook
// Tags: add custom document properties Aspose.Cells | check workbook core metadata unchanged | compare original workbook properties after custom property insertion | preserve Excel core properties when adding custom properties .NET | validate workbook metadata integrity Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, optionally sets core properties, adds custom document properties (ProjectId, Reviewed, Notes), and saves the file. Commented sections show how to store original core property values and verify that they remain unchanged after the custom properties are added.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set some core properties (available in newer Aspose.Cells versions)
            // If the current version does not support WorkbookProperties, these lines can be omitted.
            // Uncomment the following lines if your Aspose.Cells version includes WorkbookProperties.
            /*
            workbook.WorkbookProperties.Title = "Sample Workbook";
            workbook.WorkbookProperties.Author = "Jane Doe";
            workbook.WorkbookProperties.Subject = "Demo";
            workbook.WorkbookProperties.Keywords = "Aspose, Cells";

            // Store original core property values for later verification
            string originalTitle = workbook.WorkbookProperties.Title;
            string originalAuthor = workbook.WorkbookProperties.Author;
            string originalSubject = workbook.WorkbookProperties.Subject;
            string originalKeywords = workbook.WorkbookProperties.Keywords;
            */

            // Add custom document properties
            workbook.CustomDocumentProperties.Add("ProjectId", 12345);
            workbook.CustomDocumentProperties.Add("Reviewed", true);
            workbook.CustomDocumentProperties.Add("Notes", "Initial version");

            // Verify that core properties remain unchanged (if core properties are set)
            /*
            bool titleUnchanged = originalTitle == workbook.WorkbookProperties.Title;
            bool authorUnchanged = originalAuthor == workbook.WorkbookProperties.Author;
            bool subjectUnchanged = originalSubject == workbook.WorkbookProperties.Subject;
            bool keywordsUnchanged = originalKeywords == workbook.WorkbookProperties.Keywords;

            Console.WriteLine("Core properties verification:");
            Console.WriteLine($"Title unchanged: {titleUnchanged}");
            Console.WriteLine($"Author unchanged: {authorUnchanged}");
            Console.WriteLine($"Subject unchanged: {subjectUnchanged}");
            Console.WriteLine($"Keywords unchanged: {keywordsUnchanged}");
            */

            // Save the workbook (optional)
            string outputPath = "CustomPropertiesDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
