// Title: Add a custom document property named "Tags" with comma‑separated values to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a custom document property called Tags to an Excel file | store multiple tag values in Excel workbook metadata using Aspose.Cells for .NET | how to concatenate a string array into a custom property with Aspose.Cells | save workbook with custom property containing comma separated values in C# | add custom metadata to Excel workbook for content categorization with Aspose.Cells
// Tags: Aspose.Cells workbook custom metadata | Excel file tags stored as comma separated string | C# create custom property for workbook categorization | store string array in Excel custom property .NET | metadata tagging of Excel workbook using Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, builds a string array of tags, joins them into a comma‑separated string, adds this string as a custom document property named "Tags", and saves the workbook as TaggedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Define an array of tags to categorize the workbook content
        string[] tags = new string[] { "Finance", "Quarterly", "Confidential" };

        // Convert the string array to a single comma‑separated value for storage
        string tagsValue = string.Join(",", tags);

        // Add a custom document property named "Tags" containing the tags
        workbook.CustomDocumentProperties.Add("Tags", tagsValue);

        // Save the workbook (lifecycle save rule)
        workbook.Save("TaggedWorkbook.xlsx");
    }
}
