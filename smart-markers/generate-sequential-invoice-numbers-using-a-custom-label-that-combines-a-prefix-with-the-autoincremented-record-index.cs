using System;
using Aspose.Cells;

namespace InvoiceNumberGenerator
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header for the invoice number column
            sheet.Cells["A1"].PutValue("InvoiceNumber");

            // Define the custom prefix
            const string prefix = "INV-";

            // Generate sequential invoice numbers with the custom label
            // Here we manually combine the prefix with the auto‑incremented index
            for (int i = 1; i <= 10; i++)
            {
                // Cell address for the current row (column A, rows start at 2)
                string cellAddress = $"A{i + 1}";
                // Combine prefix and index to form the invoice number
                string invoiceNumber = $"{prefix}{i}";
                sheet.Cells[cellAddress].PutValue(invoiceNumber);
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save("InvoiceNumbers.xlsx", SaveFormat.Xlsx);
        }
    }
}