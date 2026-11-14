// Title: Create sequential invoice numbers with a custom prefix and zero‑padded index using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# generate invoice numbers with prefix and leading zeros | How to add sequential invoice IDs to an Excel column using Aspose.Cells | C# write custom prefixed invoice numbers to an Excel workbook with Aspose | Generate padded invoice numbers programmatically with Aspose.Cells .NET | Create invoice number series like INV-001 in Excel using Aspose.Cells
// Tags: prefixed invoice number series Aspose.Cells | zero‑padded sequence generation Excel C# | write sequential IDs to worksheet column Aspose | custom invoice identifier format .NET | auto‑increment record index workbook Aspose

using System;
using Aspose.Cells;

// The example creates a new workbook, adds a header, then generates 20 invoice numbers by concatenating the "INV-" prefix with a zero‑padded index (001‑020) and writes each to column A, finally saving the file as InvoiceNumbers.xlsx.
class InvoiceNumberGenerator
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the prefix and starting index
        string prefix = "INV-";
        int startIndex = 1;
        int totalInvoices = 20; // Example: generate 20 invoice numbers

        // Write header
        sheet.Cells["A1"].PutValue("Invoice Number");

        // Generate sequential invoice numbers with padded index
        for (int i = 0; i < totalInvoices; i++)
        {
            int currentIndex = startIndex + i;
            // Pad the numeric part with leading zeros (e.g., 001, 002)
            string invoiceNumber = $"{prefix}{currentIndex:D3}";
            // Write to column A, starting from row 2
            sheet.Cells[i + 1, 0].PutValue(invoiceNumber);
        }

        // Save the workbook to a file
        workbook.Save("InvoiceNumbers.xlsx");
    }
}
