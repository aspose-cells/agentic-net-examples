// Title: C# – Remove All Signature Lines and Add a New One in an Excel Workbook with Aspose.Cells
// Description: Loads an existing workbook, iterates through every worksheet, deletes shapes that contain a SignatureLine, creates a new SignatureLine with custom signer information, places it in cell B2 of the first sheet, and saves the file as a new workbook.
// Keywords: Aspose.Cells remove signature line C# | add signature line Excel Aspose.Cells | delete existing signature lines .NET | update signer details Excel workbook | SignatureLine shape removal | Aspose.Cells workbook protection | C# Excel digital signature line
// Common Searches: How to delete all signature lines from an Excel file using Aspose.Cells | Add a custom signature line to a worksheet with Aspose.Cells .NET | Remove picture shapes that contain signature lines in a workbook | Replace existing Excel signature lines with new signer details
// Developer Intent: Remove every existing signature line from a workbook and insert a single updated signature line with specified signer data.
// Use Cases: Refresh a signed template before re‑signing it with a new approver. | Automate signature line updates in generated reports to reflect the current reviewer. | Ensure each worksheet contains only one up‑to‑date signature line prior to distribution.
// AI Prompts: Write C# code using Aspose.Cells that removes all signature lines from every worksheet and adds a new SignatureLine at a given cell with custom signer, title, email, and instructions. | Show how to safely iterate a ShapeCollection backwards to delete picture shapes that hold a SignatureLine. | Explain the purpose of each SignatureLine property (Signer, Title, Email, Instructions, AllowComments, ShowSignedDate, IsLine) before adding it to a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    // Loads an existing workbook, iterates through every worksheet, deletes shapes that contain a SignatureLine, creates a new SignatureLine with custom signer information, places it in cell B2 of the first sheet, and saves the file as a new workbook.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook
            string inputPath = "SignedWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and remove existing signature lines
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                // Iterate backwards so removal does not affect the index order
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    // Each shape can be a Picture; check if it has an associated SignatureLine
                    if (shapes[i] is Picture picture && picture.SignatureLine != null)
                    {
                        // Remove the shape containing the signature line
                        shapes.RemoveAt(i);
                    }
                }
            }

            // Create a new signature line with updated signer details
            SignatureLine newSignature = new SignatureLine
            {
                Signer = "Alice Johnson",
                Title = "Project Manager",
                Email = "alice.johnson@example.com",
                Instructions = "Please sign to approve the document.",
                AllowComments = true,
                ShowSignedDate = true,
                IsLine = true
            };

            // Add the new signature line to the first worksheet at cell B2 (row index 1, column index 1)
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Shapes.AddSignatureLine(1, 1, newSignature);

            // Save the updated workbook
            string outputPath = "UpdatedSignatureWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
