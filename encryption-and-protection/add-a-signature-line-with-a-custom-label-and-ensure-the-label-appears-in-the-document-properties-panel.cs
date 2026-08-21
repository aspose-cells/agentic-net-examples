// Title: Add a Custom‑Labeled Signature Line and Store It in Document Properties with Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a signature line at cell B2 with a custom label (e.g., "Approval"), configures signer details, adds a matching custom document property, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells signature line C# | custom label signature line | Excel custom document property | add signature line Aspose.Cells | .NET Excel signing workflow | store signature label property
// Common Searches: Aspose.Cells add signature line with custom label | How to set a custom label for a signature line in C# | Save signature line label to Excel document properties using Aspose.Cells | Create signature line and custom property in .NET Excel file | Display signature label in workbook properties panel
// Developer Intent: Insert a signature line with a user‑defined label and expose that label through a custom document property in an Excel workbook.
// Use Cases: Automated approval forms that require a visible signature line and a searchable label for compliance audits. | Generating Excel reports where downstream systems read the "SignatureLabel" property to route documents for signing. | Building a document‑assembly pipeline that adds a signed stamp and records signer metadata in workbook properties.
// AI Prompts: Generate C# code using Aspose.Cells to add a signature line with a custom label and create a matching custom document property. | Show how to read the "SignatureLabel" custom property from an existing workbook created with Aspose.Cells. | Explain how to update the label of an existing signature line and keep the custom document property in sync.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a signature line at cell B2 with a custom label (e.g., "Approval"), configures signer details, adds a matching custom document property, and saves the file as an Excel workbook.
class AddSignatureLineWithCustomLabel
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create and configure the signature line
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Approval",                 // Custom label for the signature line
            Instructions = "Please sign to approve the document.",
            IsLine = true,
            AllowComments = true,
            ShowSignedDate = true
        };

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Add a custom document property so the label appears in the properties panel
        workbook.CustomDocumentProperties.Add("SignatureLabel", "Approval");

        // Save the workbook
        workbook.Save("SignatureLineWithLabel.xlsx");
    }
}
