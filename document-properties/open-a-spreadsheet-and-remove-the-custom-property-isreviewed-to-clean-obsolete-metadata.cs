// Title: C# – Aspose.Cells: Delete the 'IsReviewed' custom document property from an Excel workbook
// Description: Loads an XLSX file with a FileStream, checks for the custom document property IsReviewed, removes it using Aspose.Cells for .NET, and saves the workbook, effectively cleaning obsolete metadata.
// Keywords: Aspose.Cells | C# | .NET | custom document property | remove | delete | IsReviewed | Excel metadata | Workbook | FileStream | metadata cleanup
// Common Searches: Aspose.Cells remove custom property IsReviewed C# | How to delete a custom document property from Excel using .NET | C# code to clean Excel metadata with Aspose.Cells | Remove obsolete custom properties from XLSX file | Delete custom document property programmatically Aspose.Cells
// Developer Intent: Delete the 'IsReviewed' custom property from an Excel workbook and persist the updated file.
// Use Cases: Strip review flags before sharing reports with clients | Automate metadata cleanup in batch‑generated financial spreadsheets | Meet compliance requirements by removing custom properties from exported data | Prepare workbooks for archival by eliminating unnecessary metadata
// AI Prompts: Generate C# code that removes multiple custom document properties in one pass with Aspose.Cells. | Show how to list all custom properties in a workbook before deleting a specific one using Aspose.Cells. | Explain exception handling for locked or read‑only Excel files when modifying custom properties with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an XLSX file with a FileStream, checks for the custom document property IsReviewed, removes it using Aspose.Cells for .NET, and saves the workbook, effectively cleaning obsolete metadata.
class RemoveCustomProperty
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook from a file stream
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // If the custom property "IsReviewed" exists, remove it
            if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
            {
                workbook.CustomDocumentProperties.Remove("IsReviewed");
            }

            // Save the workbook with the updated metadata
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
