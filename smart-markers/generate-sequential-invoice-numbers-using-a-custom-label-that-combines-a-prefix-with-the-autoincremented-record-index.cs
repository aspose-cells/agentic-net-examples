// Title: Generate sequential invoice numbers with a custom prefix in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, adds an "InvoiceNumber" column, and fills each row with a value that concatenates a specified label (e.g., "INV-") and an auto‑incremented index, then saves the file as XLSX using Aspose.Cells. | Adapt the example to accept parameters for the starting index, total invoice count, and label, and generate the corresponding invoice numbers in the worksheet.
// Common Searches: Aspose.Cells C# generate invoice numbers with custom prefix and sequential index | How to add prefixed invoice IDs to an Excel sheet using Aspose.Cells for .NET | C# example for auto‑incrementing cell values with a label in an XLSX file | Create invoice list in Excel programmatically with Aspose.Cells and custom label
// Tags: prefixed invoice number generation Aspose.Cells | auto increment cell values C# | save workbook as XLSX Aspose.Cells | populate invoice worksheet programmatically | custom label for sequential records Excel

using Aspose.Cells;
using System;

// The sample creates a new workbook, adds column headers, and populates 15 rows with invoice numbers that combine the "INV-" prefix and an incrementing index, along with sample customer names and amounts, then saves the result as Invoices.xlsx.
class GenerateInvoiceNumbers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add column headers
        sheet.Cells["A1"].PutValue("InvoiceNumber");
        sheet.Cells["B1"].PutValue("Customer");
        sheet.Cells["C1"].PutValue("Amount");

        // Define the number of invoices and the custom prefix
        int invoiceCount = 15;
        string prefix = "INV-";

        // Fill the worksheet with sequential invoice numbers
        for (int i = 0; i < invoiceCount; i++)
        {
            int row = i + 2; // Data starts from row 2
            string invoiceNumber = prefix + (i + 1); // Prefix + auto‑incremented index
            sheet.Cells[row, 0].PutValue(invoiceNumber);
            sheet.Cells[row, 1].PutValue($"Customer {i + 1}");
            sheet.Cells[row, 2].PutValue(100 + i * 20);
        }

        // Save the workbook to an XLSX file
        workbook.Save("Invoices.xlsx", SaveFormat.Xlsx);
    }
}
