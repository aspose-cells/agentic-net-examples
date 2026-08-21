// Title: Add Integer Custom Document Property 'ProjectId' to an Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to load an existing Excel file using Aspose.Cells for .NET, add a custom document property named ProjectId with an integer value, and save the workbook preserving the new metadata.
// Keywords: Aspose.Cells | custom document property | integer property | C# | .NET | Excel metadata | ProjectId | add custom property | Workbook.Save
// Common Searches: Aspose.Cells add integer custom property | C# add ProjectId custom document property Excel | How to set numeric custom property with Aspose.Cells | Save workbook after adding custom document property .NET | Create custom metadata in Excel using Aspose.Cells
// Developer Intent: Add an integer custom document property called ProjectId to an existing Excel workbook and save the file.
// Use Cases: Store a unique project identifier inside the workbook for integration with external systems | Enable version tracking and audit trails by embedding numeric metadata | Facilitate automated processing by tagging files with a database key
// AI Prompts: Write C# code with Aspose.Cells to add a string custom property 'Author' to a workbook. | Show how to check for an existing custom property before adding 'ProjectId' using Aspose.Cells. | Provide a C# example that lists all custom document properties and their values in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to load an existing Excel file using Aspose.Cells for .NET, add a custom document property named ProjectId with an integer value, and save the workbook preserving the new metadata.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom document property named "ProjectId" with an integer value
        // Uses CustomDocumentPropertyCollection.Add(string, int) overload
        workbook.CustomDocumentProperties.Add("ProjectId", 12345);

        // Save the workbook with the new custom property
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
