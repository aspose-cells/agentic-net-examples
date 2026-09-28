// Title: Add a custom Boolean document property 'IsReviewed' to an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program that loads an existing .xlsx file using Aspose.Cells, adds a custom Boolean document property named IsReviewed set to true, and saves the updated workbook. | Show how to use Aspose.Cells' CustomDocumentProperties API to insert a Boolean metadata field into an Excel workbook and persist the change.
// Common Searches: Aspose.Cells C# add custom Boolean document property to existing Excel file | How to set a custom workbook property called IsReviewed using .NET | Programmatically insert custom metadata into an .xlsx with Aspose.Cells | C# code to add and save a custom document property in an Excel workbook | Aspose.Cells example for adding Boolean custom document properties
// Tags: add custom boolean property Aspose.Cells | custom document properties Excel .NET | Aspose.Cells workbook metadata manipulation | set IsReviewed property C# | save workbook with new custom property

using Aspose.Cells;
using System;

// Loads 'input.xlsx', adds a custom Boolean document property named 'IsReviewed' set to true via Aspose.Cells, and saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom Boolean property named "IsReviewed" and set it to true
        workbook.CustomDocumentProperties.Add("IsReviewed", true);

        // Save the workbook with the new property
        workbook.Save("output.xlsx");
    }
}
