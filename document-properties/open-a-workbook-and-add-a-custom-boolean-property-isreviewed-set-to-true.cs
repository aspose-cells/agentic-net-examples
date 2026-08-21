// Title: C# – Add a Boolean Custom Document Property "IsReviewed" to an Aspose.Cells Workbook and Save as XLSX
// Description: Demonstrates how to create a new Workbook with Aspose.Cells for .NET, add a custom Boolean document property named IsReviewed set to true, and persist the file as ReviewedWorkbook.xlsx in XLSX format.
// Keywords: Aspose.Cells C# custom document property | add Boolean property IsReviewed | save workbook with custom properties | Aspose.Cells API custom property example | Excel metadata Boolean flag | C# Aspose.Cells create workbook | XLSX custom property Aspose
// Common Searches: Aspose.Cells add custom Boolean property C# | How to set IsReviewed document property in Excel using Aspose | Save workbook with custom properties Aspose.Cells .NET | Create Excel file with metadata flag using Aspose.Cells
// Developer Intent: Create a workbook, attach a Boolean custom property called IsReviewed with value true, and write the file to disk.
// Use Cases: Flag a spreadsheet as reviewed before sharing it with stakeholders. | Store processing status inside the file for downstream automation pipelines. | Embed custom audit information directly within an Excel document.
// AI Prompts: Generate C# code that adds a Boolean custom document property "IsReviewed" to an existing Aspose.Cells workbook and saves it. | Show how to read the "IsReviewed" custom property from a workbook using Aspose.Cells for .NET. | List all custom document properties in a workbook with Aspose.Cells and display their values.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyDemo
{
    // Demonstrates how to create a new Workbook with Aspose.Cells for .NET, add a custom Boolean document property named IsReviewed set to true, and persist the file as ReviewedWorkbook.xlsx in XLSX format.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add a custom Boolean property named "IsReviewed" with value true
            workbook.CustomDocumentProperties.Add("IsReviewed", true);

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("ReviewedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
