// Title: C# – Remove All Signature Lines and Insert a New One in an Excel Workbook with Aspose.Cells
// Description: Load an Excel file with Aspose.Cells for .NET, iterate through each worksheet, delete every picture shape that contains a SignatureLine, create a new SignatureLine with custom signer information, place it at a specified cell, and save the updated workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel | SignatureLine | remove signature line | add signature line | update signer details | ShapeCollection | Picture | automation
// Common Searches: How to delete signature lines in Excel using Aspose.Cells C# | Add a signature line to a specific cell with Aspose.Cells .NET | Replace existing signature lines programmatically | Remove all picture signatures from a workbook | Update signer name, title, and email in Excel signature line
// Developer Intent: Programmatically clear all existing signature line objects from a workbook and add a fresh signature line with defined signer attributes.
// Use Cases: Clean outdated signature lines before re‑signing a contract workbook. | Batch‑update signer name, title, and email across multiple worksheets. | Prepare a template by removing placeholder signatures and inserting a new approver line.
// AI Prompts: Generate C# code using Aspose.Cells that removes every signature line picture from an Excel file and adds a new SignatureLine with custom signer name, title, email, and instructions. | Explain how to safely iterate a worksheet's ShapeCollection, detect Picture objects with a non‑null SignatureLine, and delete them without index errors. | Provide troubleshooting steps if the new signature line does not appear after calling shapes.AddSignatureLine.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel file with Aspose.Cells for .NET, iterate through each worksheet, delete every picture shape that contains a SignatureLine, create a new SignatureLine with custom signer information, place it at a specified cell, and save the updated workbook.
class SignatureLineUpdater
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            ShapeCollection shapes = sheet.Shapes;

            // Identify indices of all signature line pictures
            List<int> indicesToRemove = new List<int>();
            for (int i = 0; i < shapes.Count; i++)
            {
                // Signature lines are stored as Picture objects with a non‑null SignatureLine
                if (shapes[i] is Picture pic && pic.SignatureLine != null)
                {
                    indicesToRemove.Add(i);
                }
            }

            // Remove identified signature lines (remove from highest index to keep collection stable)
            for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                shapes.RemoveAt(indicesToRemove[i]);
            }

            // Create a new signature line with updated signer details
            SignatureLine newSignature = new SignatureLine
            {
                Signer = "Jane Doe",
                Title = "Project Manager",
                Email = "jane.doe@example.com",
                Instructions = "Please sign to approve.",
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true
            };

            // Add the new signature line at the desired cell position (row 5, column 2)
            int topRow = 5;      // zero‑based row index
            int leftColumn = 2;  // zero‑based column index
            shapes.AddSignatureLine(topRow, leftColumn, newSignature);
        }

        // Save the workbook with the updated signature line
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
