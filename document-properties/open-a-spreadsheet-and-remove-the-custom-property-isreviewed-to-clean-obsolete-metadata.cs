// Title: C# – Remove the “IsReviewed” custom document property from an Excel workbook using Aspose.Cells for .NET
// Description: This example opens an XLSX file with a FileStream, checks the Workbook.CustomDocumentProperties collection for a property named "IsReviewed", deletes it when present, and saves the cleaned workbook as a new XLSX file. It demonstrates safe metadata cleanup with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel | custom document property | remove property | IsReviewed | metadata cleanup | Workbook.CustomDocumentProperties | SaveFormat.Xlsx | file stream | Excel automation | document properties API
// Common Searches: Aspose.Cells remove custom property IsReviewed | C# delete Excel custom document property | How to clear obsolete metadata in XLSX using Aspose.Cells | Check and remove custom document property in .NET | Batch remove custom properties from Excel files C#
// Developer Intent: Delete the "IsReviewed" custom document property from an Excel workbook and save the updated file.
// Use Cases: Strip review flags before distributing reports to external partners. | Automate metadata sanitization across a batch of generated workbooks. | Ensure compliance by removing proprietary custom properties prior to archiving.
// AI Prompts: Generate C# code with Aspose.Cells that removes the custom document property "IsReviewed" from an Excel file and saves the result. | Explain how to verify the existence of a custom property and delete it safely using Aspose.Cells for .NET. | Show a script that processes all XLSX files in a folder, removing the "IsReviewed" property from each workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This example opens an XLSX file with a FileStream, checks the Workbook.CustomDocumentProperties collection for a property named "IsReviewed", deletes it when present, and saves the cleaned workbook as a new XLSX file. It demonstrates safe metadata cleanup with Aspose.Cells for .NET.
class RemoveCustomProperty
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Open the workbook using a file stream (load rule)
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // Remove the custom property "IsReviewed" if it exists
            if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
            {
                workbook.CustomDocumentProperties.Remove("IsReviewed");
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
