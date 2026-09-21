// Title: Insert a clickable signature line with hyperlink to a professional profile in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a textbox shape to a worksheet, fills it with custom signature text, and attaches a hyperlink to the signer's professional profile with a specific display label using Aspose.Cells. | Modify an existing Aspose.Cells workbook to include a signature placeholder that links to a URL, ensuring the hyperlink’s visible text shows the signer’s name.
// Common Searches: how to add a hyperlink to a textbox shape in Excel with Aspose.Cells C# | Aspose.Cells create signature line with clickable URL in .NET | set hyperlink display text for a shape in an Excel file using C# | programmatically embed LinkedIn profile link in Excel signature placeholder Aspose.Cells
// Tags: Aspose.Cells add hyperlink to textbox shape | C# create Excel signature line with URL | set shape hyperlink display text Aspose.Cells | Excel workbook signature placeholder hyperlink | Aspose.Cells textbox shape formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureExample
{
    // The example creates a new workbook, adds a textbox shape spanning cells B2‑F6 as a signature placeholder, populates it with signer details, attaches a clickable hyperlink to the signer's professional profile with custom display text, ensures the output directory exists, and saves the file as SignedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the position for the signature placeholder (cells B2 to F6 as an example)
                int upperLeftRow = 1;      // B2 row (zero‑based)
                int upperLeftColumn = 1;   // B2 column (zero‑based)
                int lowerRightRow = 5;     // end row
                int lowerRightColumn = 5;  // end column

                // Width and height of the textbox in pixels (required parameters)
                int width = 200;
                int height = 80;

                // Information to display in the signature placeholder
                string signer = "John Doe";
                string signerTitle = "Senior Analyst";
                string email = "john.doe@example.com";

                // Add a textbox shape to act as a signature line placeholder
                Shape signatureBox = sheet.Shapes.AddTextBox(
                    upperLeftRow,
                    upperLeftColumn,
                    lowerRightRow,
                    lowerRightColumn,
                    width,
                    height);

                // Set the displayed text
                signatureBox.Text = $"Signed by: {signer}\nTitle: {signerTitle}\nEmail: {email}";
                signatureBox.Font.IsBold = true;
                signatureBox.Font.Size = 12;
                signatureBox.Line.Weight = 1.0; // optional border

                // Define output file path
                string outputPath = "SignedWorkbook.xlsx";

                // Ensure the output directory exists (handle case when only file name is provided)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the added signature placeholder
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the signed workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
