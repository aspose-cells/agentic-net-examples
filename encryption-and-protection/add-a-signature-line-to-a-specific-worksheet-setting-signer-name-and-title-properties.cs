// Title: Add a Signature Line with Signer Name & Title to an Excel Worksheet using Aspose.Cells for .NET
// Description: Shows how to create a Workbook, select a Worksheet, configure a SignatureLine (signer, title, email, instructions, display options) and insert it at a specific cell via Shapes.AddSignatureLine, then save the workbook.
// Keywords: Aspose.Cells | C# | .NET | SignatureLine | add signature line | signer name | signer title | Excel worksheet | digital signature | Shapes.AddSignatureLine | office automation | document approval | programmatic Excel
// Common Searches: Aspose.Cells add signature line C# | How to set signer name in Excel using Aspose.Cells | SignatureLine Title property Aspose.Cells .NET | Insert digital signature line into worksheet programmatically | C# code for Aspose.Cells signature line | Add approval line to Excel file Aspose
// Developer Intent: Insert a signature line into a worksheet and define signer name and title.
// Use Cases: Generate reports that require manager approval directly in the Excel file. | Create contract templates with pre‑populated signature placeholders for multiple parties. | Automate compliance documentation by adding signed‑date lines to generated spreadsheets. | Build a workflow where each department adds its own signature line to a shared workbook.
// AI Prompts: Write C# code using Aspose.Cells to place a signature line with signer 'Emily Chen' and title 'Finance Director' at cell D5. | Explain how to update the Signer and Title of an existing SignatureLine object in a saved workbook. | Show how to add three different signature lines to three worksheets, each with unique signer details and custom instructions. | Provide a step‑by‑step guide for adding a signature line that includes email and comment permissions using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    // Shows how to create a Workbook, select a Worksheet, configure a SignatureLine (signer, title, email, instructions, display options) and insert it at a specific cell via Shapes.AddSignatureLine, then save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a SignatureLine object and set signer properties
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",          // Signer's name
                Title = "Chief Technology Officer", // Signer's title
                Email = "john.doe@example.com", // Optional: email
                IsLine = true,                // Display as a line
                AllowComments = true,         // Allow comments
                ShowSignedDate = true,        // Show signed date
                Instructions = "Please sign to approve."
            };

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            // This uses the ShapeCollection.AddSignatureLine method.
            worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("SignatureLineDemo.xlsx");
        }
    }
}
