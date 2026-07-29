// Title: Copy rows with hyperlinks and verify URLs using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a worksheet row that contains a hyperlink, preserve the link with CopyOptions.ExtendToAdjacentRange, and programmatically confirm the copied URL matches the original.
// Keywords: Aspose.Cells copy row hyperlink | C# Aspose.Cells hyperlink verification | CopyOptions ExtendToAdjacentRange example | preserve hyperlink when copying rows | Aspose.Cells .NET hyperlink address check
// Common Searches: Aspose.Cells copy row with hyperlink C# | verify hyperlink after copying rows Aspose.Cells | CopyOptions ExtendToAdjacentRange hyperlink | how to keep hyperlink address when duplicating rows | C# Aspose.Cells copy rows preserving links
// Developer Intent: Duplicate a row that includes a hyperlink and ensure the new row points to the same URL.
// Use Cases: Create templated report rows where each copy retains its original web link. | Migrate data rows containing external references to another sheet without breaking navigation. | Automate generation of catalog entries that need identical hyperlink destinations across rows.
// AI Prompts: Write C# code with Aspose.Cells to copy a row containing a hyperlink and validate that the copied link URL is identical to the source. | Show how to use CopyOptions.ExtendToAdjacentRange to duplicate a row with hyperlinks and then iterate Worksheet.Hyperlinks to confirm address consistency. | Explain handling multiple hyperlinks in a copied row and how to programmatically verify each copied URL matches its source.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkCopyDemo
{
    // Demonstrates how to copy a worksheet row that contains a hyperlink, preserve the link with CopyOptions.ExtendToAdjacentRange, and programmatically confirm the copied URL matches the original.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data in row 2 (index 1) and a hyperlink in cell B2
            sheet.Cells["A2"].PutValue("Original Row");
            sheet.Hyperlinks.Add("B2", 1, 1, "https://www.example.com");

            // Store original hyperlink address for later verification
            string originalAddress = sheet.Hyperlinks[0].Address;

            // Prepare copy options: extend the hyperlink range to the adjacent row
            CopyOptions options = new CopyOptions();
            options.ExtendToAdjacentRange = true;

            // Copy row 2 (index 1) to row 3 (index 2) using the copy options
            sheet.Cells.CopyRows(sheet.Cells, 1, 2, 1, options);

            // Verify that a hyperlink now exists in the destination row (B3) and
            // that it points to the same address as the original hyperlink
            bool hyperlinkFound = false;
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                // Check if the hyperlink is located in the destination row (row index 2)
                if (link.Area.StartRow == 2 && link.Area.StartColumn == 1) // column B = index 1
                {
                    hyperlinkFound = true;
                    Console.WriteLine("Hyperlink copied to B3.");
                    Console.WriteLine("Original address: " + originalAddress);
                    Console.WriteLine("Copied address  : " + link.Address);
                    Console.WriteLine("Address match   : " + (originalAddress == link.Address));
                    break;
                }
            }

            if (!hyperlinkFound)
            {
                Console.WriteLine("Hyperlink was not copied to the destination row.");
            }

            // Save the workbook to verify the result manually if needed
            workbook.Save("HyperlinkCopyResult.xlsx");
        }
    }
}
