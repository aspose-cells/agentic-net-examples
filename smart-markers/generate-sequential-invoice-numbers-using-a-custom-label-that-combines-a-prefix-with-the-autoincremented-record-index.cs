// Title: Create sequential invoice numbers with a custom prefix in Excel using Aspose.Cells for .NET
// Description: This example shows how to build a new workbook, add an "InvoiceNumber" header, and fill column A with 20 identifiers that combine a fixed prefix (e.g., "INV-") and an auto‑incremented index starting at 1. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells invoice numbers C# | custom prefix Excel IDs | auto increment column values .NET | generate sequential IDs Aspose | save workbook as xlsx C#
// Common Searches: Aspose.Cells generate sequential IDs with prefix | C# create invoice list in Excel | how to add custom prefix to Excel column using Aspose | auto‑increment numbers in Aspose.Cells workbook | save generated invoice numbers to XLSX in .NET
// Developer Intent: Produce an Excel sheet that lists prefixed, sequential invoice numbers.
// Use Cases: Pre‑populate a batch of invoice numbers for a billing run before adding line items. | Create a printable invoice register where each row already contains a unique ID. | Integrate automatic numbering into a larger invoice‑generation pipeline that writes other details later.
// AI Prompts: Generate C# code with Aspose.Cells that writes prefixed sequential invoice numbers to an XLSX file. | Provide a reusable method that returns a list of invoice numbers given a prefix and count, using Aspose.Cells. | Explain how to start numbering from a custom offset and optionally append a date segment to each invoice ID.

using System;
using Aspose.Cells;

namespace InvoiceNumberGenerator
{
    // This example shows how to build a new workbook, add an "InvoiceNumber" header, and fill column A with 20 identifiers that combine a fixed prefix (e.g., "INV-") and an auto‑incremented index starting at 1. The workbook is then saved as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a custom prefix for invoice numbers
            const string prefix = "INV-";

            // Add header for the invoice number column
            sheet.Cells["A1"].PutValue("InvoiceNumber");

            // Number of invoices to generate
            int invoiceCount = 20;

            // Populate sequential invoice numbers with the custom prefix
            for (int i = 0; i < invoiceCount; i++)
            {
                // Row index in Aspose.Cells is zero‑based; +1 for header row
                int rowIndex = i + 1;
                // Combine prefix with the auto‑incremented index (starting at 1)
                string invoiceNumber = $"{prefix}{i + 1}";
                sheet.Cells[rowIndex, 0].PutValue(invoiceNumber);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Invoices.xlsx", SaveFormat.Xlsx);
        }
    }
}
