// Title: Add a Hyperlinked Signature Line with Custom Display Text in Excel using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, configure a SignatureLine with signer details, embed a clickable URL to the signer's professional profile, set a custom link label, place the signature line in a worksheet cell, and save the file as an XLSX document.
// Keywords: Aspose.Cells signature line hyperlink | C# Excel signature line custom text | add clickable profile link to Excel signature | Aspose.Cells add signature line with URL | hyperlinked signature line .NET | Excel digital signature link | US developers Aspose.Cells example
// Common Searches: How to add a URL to a signature line in Excel with Aspose.Cells C# | Set custom display text for a signature line hyperlink using Aspose.Cells | Aspose.Cells example for hyperlinked signature line | C# code to embed LinkedIn link in Excel signature line
// Developer Intent: Insert a signature line into an Excel worksheet that contains a clickable link to the signer’s online profile and shows a custom label.
// Use Cases: Contracts where each signer’s name links to their LinkedIn or company bio. | Automated reports that require a signed section with a quick‑access profile page. | Internal approval sheets that let reviewers open a signer’s contact page directly from the workbook.
// AI Prompts: Generate C# code with Aspose.Cells to add a signature line at cell C3 that links to https://linkedin.com/in/johndoe and displays the text 'John Doe – LinkedIn'. | Explain how to modify SignatureLine properties to include a clickable URL and custom link text in an Excel file using Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for creating a workbook, inserting a hyperlinked signature line, and saving it as an XLSX file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineExample
{
    // Demonstrates how to create a new Workbook, configure a SignatureLine with signer details, embed a clickable URL to the signer's professional profile, set a custom link label, place the signature line in a worksheet cell, and save the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure the signature line properties
                SignatureLine signatureLine = new SignatureLine
                {
                    Signer = "John Doe",                     // Display name of the signer
                    Title = "Software Engineer",            // Signer's title
                    Email = "john.doe@example.com",         // Signer's email (optional)
                    IsLine = true,                          // Indicates this is a line
                    AllowComments = true,                   // Allow comments on the signature
                    ShowSignedDate = true,                  // Show the date when signed
                    Instructions = "Professional profile: https://linkedin.com/in/johndoe"
                };

                // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
                Picture signaturePicture = worksheet.Shapes.AddSignatureLine(2, 2, signatureLine);

                // Define output file path
                string outputPath = "SignatureLineWithHyperlink.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
