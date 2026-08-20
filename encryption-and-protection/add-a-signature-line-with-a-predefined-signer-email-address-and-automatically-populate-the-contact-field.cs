// Title: Insert a Signature Line with Pre‑filled Signer Email in Excel using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, add a SignatureLine shape, set the signer name, title, and email (contact field), configure visual options such as line display and signed date, place it at cell B2, and save the file as SignatureLineWithEmail.xlsx.
// Keywords: Aspose.Cells signature line C# | Excel add signature line programmatically | predefined signer email Aspose.Cells | populate contact field SignatureLine | digital signature line .NET | Excel workbook protection Aspose
// Common Searches: how to add a signature line with email using Aspose.Cells | Aspose.Cells set signer email for SignatureLine | C# insert digital signature line in Excel | populate contact field of SignatureLine programmatically | Aspose.Cells add signature line to specific cell
// Developer Intent: Add a SignatureLine to an Excel worksheet and automatically fill the signer’s email address.
// Use Cases: Generate financial statements that include a pre‑filled approval signature line. | Create contract templates with a ready‑to‑sign line for automated workflow routing. | Embed a compliance‑required signature line that records the signed date in audit reports.
// AI Prompts: Write C# code with Aspose.Cells to place a SignatureLine at B2, setting signer name, title, and email. | Show how to customize the visual style of a SignatureLine (line only, show signed date) and save the workbook. | Explain how to read, update, or remove the Email property of an existing SignatureLine in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    // Demonstrates how to create a new Workbook, add a SignatureLine shape, set the signer name, title, and email (contact field), configure visual options such as line display and signed date, place it at cell B2, and save the file as SignatureLineWithEmail.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a SignatureLine object
            SignatureLine signatureLine = new SignatureLine();

            // Set predefined signer information
            signatureLine.Signer = "John Doe";
            signatureLine.Title = "Approver";
            signatureLine.Email = "john.doe@example.com"; // contact field

            // Optional visual settings
            signatureLine.IsLine = true;          // display as a line
            signatureLine.ShowSignedDate = true; // show date after signing

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Save the workbook (lifecycle: save)
            workbook.Save("SignatureLineWithEmail.xlsx");
        }
    }
}
