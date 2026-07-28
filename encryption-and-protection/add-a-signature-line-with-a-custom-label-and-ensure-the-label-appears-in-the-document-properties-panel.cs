// Title: C# – Add a Signature Line with Custom Label and Save It as a Document Property using Aspose.Cells
// Description: Demonstrates how to create a new Workbook, insert a SignatureLine with a custom Instructions label, add the same label as a custom document property, and save the file as an .xlsx workbook. The custom label becomes visible in Excel's document‑properties panel, enabling easy searching and filtering.
// Keywords: Aspose.Cells | C# signature line | custom label | document property | Excel | SignatureLine | CustomDocumentProperties | .NET | add signature line | custom instructions
// Common Searches: Aspose.Cells add signature line C# | custom label for signature line Excel | store signature instructions as document property | how to show signature line text in Excel properties | C# example for SignatureLine with custom instructions
// Developer Intent: Insert a SignatureLine with a user‑defined label and expose that label through a custom document property in the saved workbook.
// Use Cases: Create a template that requires a specific signer and displays custom signing instructions. | Expose signing instructions in the workbook's properties for automated document management. | Enable downstream processes to locate Excel files based on the signature label via property filters.
// AI Prompts: Show C# code that adds a SignatureLine with a custom label and writes the label to a custom document property using Aspose.Cells. | Give an example of reading the "SignatureLabel" property from an existing workbook. | Explain how to modify the Instructions text of a SignatureLine and update the corresponding custom document property.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new Workbook, insert a SignatureLine with a custom Instructions label, add the same label as a custom document property, and save the file as an .xlsx workbook. The custom label becomes visible in Excel's document‑properties panel, enabling easy searching and filtering.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure the signature line with a custom label (Instructions)
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Manager",
                Email = "john.doe@example.com",
                Instructions = "Please sign the document", // custom label
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true
            };

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Add a custom document property so the label appears in the properties panel
            workbook.CustomDocumentProperties.Add("SignatureLabel", signatureLine.Instructions);

            // Save the workbook
            workbook.Save("SignatureLineWithCustomLabel.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
