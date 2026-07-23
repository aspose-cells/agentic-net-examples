// Title: Add Integer Custom Document Property "ProjectId" to an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Loads input.xlsx, adds a custom document property named ProjectId with the integer value 12345, and saves the modified file as output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | custom document property | integer property | ProjectId | Excel metadata | add custom property | save workbook
// Common Searches: add integer custom property Aspose.Cells C# | set ProjectId Excel metadata .NET | how to add custom document property to existing workbook | Aspose.Cells save workbook after adding property
// Developer Intent: Create and persist an integer custom document property named ProjectId in an existing Excel file.
// Use Cases: Link workbooks to project tracking systems | Store audit identifiers for compliance reporting | Enable metadata‑driven filtering in document repositories
// AI Prompts: Write C# code with Aspose.Cells to add a custom string property "Author" to an Excel file and save it. | Show how to read the ProjectId custom property from a workbook and handle the case where the property is missing. | Demonstrate updating an existing custom document property value using Aspose.Cells in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads input.xlsx, adds a custom document property named ProjectId with the integer value 12345, and saves the modified file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom document property named "ProjectId" with an integer value
        workbook.CustomDocumentProperties.Add("ProjectId", 12345);

        // Save the workbook with the new property
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
