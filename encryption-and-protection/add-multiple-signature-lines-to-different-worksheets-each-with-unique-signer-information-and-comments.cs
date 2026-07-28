// Title: Add Multiple Signature Lines with Unique Signer Details to Separate Worksheets using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add Finance, HR, and IT sheets, configure three SignatureLine objects with distinct signer, title, email, instructions, comment and date settings, insert each line into a specific cell via Shapes.AddSignatureLine, and save the file as MultipleSignatureLines.xlsx.
// Keywords: Aspose.Cells signature line C# | multiple signature lines Excel | different worksheets signature line | custom signer information | AllowComments property | ShowSignedDate option | Shapes.AddSignatureLine | digital signature workbook .NET | Excel protection Aspose.Cells
// Common Searches: aspnet add signature line to multiple worksheets | unique signer info per signature line Aspose.Cells | place signature line in specific cell C# | configure AllowComments and ShowSignedDate for Excel signature lines | save workbook with several signature lines Aspose
// Developer Intent: Generate a .NET workbook that contains separate signature lines on different worksheets, each populated with its own signer name, title, email, instructions, and comment/date preferences.
// Use Cases: Finance sheet: CFO signature line to approve financial statements. | HR sheet: HR manager signature line for onboarding confirmation. | IT sheet: IT director signature line to validate system upgrade approval.
// AI Prompts: Write C# code using Aspose.Cells that adds a signature line with custom signer name, title, email, and instructions to cell C3. | Show how to insert multiple signature lines with varying AllowComments and ShowSignedDate settings across several worksheets in one workbook. | Provide an example that saves an Excel file containing signature lines on three different sheets, each with unique signer details.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLinesDemo
{
    // Shows how to create a workbook, add Finance, HR, and IT sheets, configure three SignatureLine objects with distinct signer, title, email, instructions, comment and date settings, insert each line into a specific cell via Shapes.AddSignatureLine, and save the file as MultipleSignatureLines.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Worksheet 1 - add a signature line with signer A
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Finance";

            // Configure signature line for signer A
            SignatureLine sigLine1 = new SignatureLine
            {
                Signer = "Alice Johnson",
                Title = "Chief Financial Officer",
                Email = "alice.johnson@example.com",
                Instructions = "Please sign to approve the financial report.",
                AllowComments = true,
                ShowSignedDate = true,
                IsLine = true,
                Id = Guid.NewGuid(),
                ProviderId = Guid.NewGuid()
            };

            // Add the signature line to cell B2 (row index 1, column index 1)
            sheet1.Shapes.AddSignatureLine(1, 1, sigLine1);

            // -------------------------------------------------
            // Worksheet 2 - add a signature line with signer B
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("HR");
            // Configure signature line for signer B
            SignatureLine sigLine2 = new SignatureLine
            {
                Signer = "Bob Smith",
                Title = "HR Manager",
                Email = "bob.smith@example.com",
                Instructions = "Sign to confirm employee onboarding.",
                AllowComments = false,
                ShowSignedDate = true,
                IsLine = true,
                Id = Guid.NewGuid(),
                ProviderId = Guid.NewGuid()
            };

            // Add the signature line to cell D5 (row index 4, column index 3)
            sheet2.Shapes.AddSignatureLine(4, 3, sigLine2);

            // -------------------------------------------------
            // Worksheet 3 - add a signature line with signer C
            // -------------------------------------------------
            Worksheet sheet3 = workbook.Worksheets.Add("IT");
            // Configure signature line for signer C
            SignatureLine sigLine3 = new SignatureLine
            {
                Signer = "Carol Lee",
                Title = "IT Director",
                Email = "carol.lee@example.com",
                Instructions = "Approve the system upgrade.",
                AllowComments = true,
                ShowSignedDate = false,
                IsLine = true,
                Id = Guid.NewGuid(),
                ProviderId = Guid.NewGuid()
            };

            // Add the signature line to cell A10 (row index 9, column index 0)
            sheet3.Shapes.AddSignatureLine(9, 0, sigLine3);

            // Save the workbook with all signature lines
            workbook.Save("MultipleSignatureLines.xlsx");
        }
    }
}
