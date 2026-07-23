// Title: C# – Add a 'Tags' custom document property (comma‑separated) to an Aspose.Cells workbook
// Description: This .NET example creates a new Workbook, builds a string array of tags (e.g., Finance, 2023, Quarterly), joins them into a comma‑separated value, stores it as a custom document property named "Tags", and saves the file as WorkbookWithTags.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | custom document property | Tags property | comma separated tags | Excel metadata | store tags in workbook | Workbook custom properties | Aspose.Cells example | Excel file tagging | document properties API | GitHub sample | code snippet
// Common Searches: how to add a custom document property in Aspose.Cells C# | store multiple tags in an Excel workbook using Aspose.Cells | Aspose.Cells add Tags property to workbook | C# example for custom properties in Excel with Aspose | comma separated custom property Aspose.Cells .NET
// Developer Intent: Create and assign a custom document property called "Tags" that holds a list of tag strings in an Aspose.Cells workbook.
// Use Cases: Label financial reports with categories such as Finance, 2023, Quarterly for easy filtering. | Enable automated scripts to locate workbooks by reading the "Tags" property. | Persist user‑defined classifications within the Excel file for downstream analytics.
// AI Prompts: Show how to read the "Tags" custom property from a workbook and split it back into a string array using Aspose.Cells. | Demonstrate updating the "Tags" property by appending a new tag without overwriting existing values. | Provide a C# example that adds multiple custom document properties of different data types (string, date, number) to a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This .NET example creates a new Workbook, builds a string array of tags (e.g., Finance, 2023, Quarterly), joins them into a comma‑separated value, stores it as a custom document property named "Tags", and saves the file as WorkbookWithTags.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define an array of tags to categorize the workbook content
        string[] tags = new string[] { "Finance", "2023", "Quarterly" };

        // Convert the string array to a single string (comma‑separated) for storage
        string tagsValue = string.Join(",", tags);

        // Add a custom document property named "Tags" with the concatenated tag string
        workbook.CustomDocumentProperties.Add("Tags", tagsValue);

        // Save the workbook to a file
        workbook.Save("WorkbookWithTags.xlsx");
    }
}
