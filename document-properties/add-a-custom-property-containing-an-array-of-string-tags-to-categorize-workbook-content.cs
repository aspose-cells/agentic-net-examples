// Title: Add a comma‑separated "Tags" custom document property to an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add custom document property with multiple values | store a list of tags in an Excel file using Aspose.Cells custom properties | how to concatenate string array into a single property in Aspose.Cells workbook | save custom metadata called Tags in an .xlsx using Aspose.Cells for .NET | example code for adding custom document property to a new workbook in C#
// Tags: add custom document property Aspose.Cells C# | comma separated metadata tags Excel workbook | custom workbook property array workaround Aspose.Cells | Aspose.Cells example storing tags in custom properties | Excel file custom metadata using Aspose.Cells API | C# combine string array for workbook property

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, joins a string array of tags into a comma‑separated value, adds it as a custom document property named "Tags" via workbook.CustomDocumentProperties.Add, ensures the output directory exists, and saves the file as TaggedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define the tags to be stored as a custom property
            string[] tags = new string[] { "Finance", "Q1", "Report" };

            // Combine tags into a single string (Aspose.Cells does not support array values directly)
            string tagsValue = string.Join(",", tags);

            // Add a custom document property named "Tags" containing the combined string
            // Overload: Add(string name, object value)
            workbook.CustomDocumentProperties.Add("Tags", tagsValue);

            // Define output file path
            string outputPath = "TaggedWorkbook.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
