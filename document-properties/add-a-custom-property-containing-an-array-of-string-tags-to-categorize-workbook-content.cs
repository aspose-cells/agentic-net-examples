// Title: C# – Add a 'Tags' custom document property (comma‑separated) to an Aspose.Cells workbook
// Description: Demonstrates how to create a new Workbook with Aspose.Cells, convert a string[] of tags into a single comma‑separated value, store it in a custom document property named "Tags", and save the file as WorkbookWithTags.xlsx.
// Keywords: Aspose.Cells C# custom property | Excel workbook tags | add custom document property Aspose | store array as string Excel | metadata tagging Aspose.Cells | comma separated values Excel | C# .NET Excel metadata | GitHub sample Aspose.Cells | global Excel tagging example
// Common Searches: Aspose.Cells add Tags property C# | store multiple tags in Excel workbook using Aspose | custom document property array Aspose.Cells .NET | how to save tag list in Excel with Aspose.Cells | C# example for Excel metadata tagging
// Developer Intent: Insert a custom document property called "Tags" that holds a list of tag strings for workbook classification.
// Use Cases: Label financial statements with tags like Finance, Report, 2023 for easy cataloging. | Enable downstream services to filter Excel files based on tag metadata. | Persist project, department, or version identifiers within the workbook for audit trails.
// AI Prompts: Write C# code using Aspose.Cells to store tags as a JSON array in a custom property instead of a CSV string. | Show how to read the "Tags" property from a workbook and convert it back to a string[] in C#. | Provide robust error‑handling patterns when adding or updating custom document properties with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTagExample
{
    // Demonstrates how to create a new Workbook with Aspose.Cells, convert a string[] of tags into a single comma‑separated value, store it in a custom document property named "Tags", and save the file as WorkbookWithTags.xlsx.
    public class AddTagProperty
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define an array of tags
                string[] tags = new string[] { "Finance", "Report", "2023" };

                // Convert the array to a single string (comma‑separated)
                string tagsValue = string.Join(",", tags);

                // Add a custom document property named "Tags" with the concatenated string value
                workbook.CustomDocumentProperties.Add("Tags", tagsValue);

                // Save the workbook to a file
                workbook.Save("WorkbookWithTags.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddTagProperty.Run();
        }
    }
}
