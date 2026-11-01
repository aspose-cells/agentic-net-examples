// Title: Generate sequential invoice numbers with a custom prefix in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write a C# method that uses Aspose.Cells to insert an 'Invoice Number' column where each value combines a configurable prefix and a zero‑padded incremental index. | Refactor the sample to accept the prefix string and the starting index as method arguments and return the populated Workbook object. | Add logic to auto‑fit all columns, save the workbook to a given .xlsx path, and print the full file path to the console after the export.
// Common Searches: asp.net aspose.cells generate invoice numbers with custom prefix and zero padding | c# create sequential invoice IDs in Excel using Aspose.Cells library | how to add auto‑incrementing invoice column with prefix in Aspose.Cells workbook | aspose.cells c# zero‑pad sequential numbers for invoice report
// Tags: generate sequential invoice numbers Aspose.Cells | custom prefix invoice ID Excel C# | zero‑padded numbering Aspose.Cells workbook | auto‑fit columns Aspose.Cells export | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// Creates an Excel workbook (InvoiceReport.xlsx) with a header row and sample customer data, adding an 'Invoice Number' column where each entry is the fixed "INV-" prefix followed by a zero‑padded sequential index starting at 1, then auto‑fits columns and saves the file.
class InvoiceNumberGenerator
{
    static void Main()
    {
        // Define the prefix for invoice numbers
        const string invoicePrefix = "INV-";

        // Starting index for the first invoice (can be changed as needed)
        int startIndex = 1;

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Invoices";

        // Add header row
        sheet.Cells["A1"].PutValue("Invoice Number");
        sheet.Cells["B1"].PutValue("Customer");
        sheet.Cells["C1"].PutValue("Amount");

        // Sample data rows (replace with actual data source as needed)
        var customers = new[]
        {
            new { Name = "Acme Corp", Amount = 1250.00 },
            new { Name = "Beta Ltd", Amount = 980.50 },
            new { Name = "Gamma Inc", Amount = 4320.75 }
        };

        // Populate data and generate sequential invoice numbers
        for (int i = 0; i < customers.Length; i++)
        {
            // Row index in Excel (1‑based, plus header row)
            int excelRow = i + 2;

            // Generate invoice number: prefix + zero‑padded index
            string invoiceNumber = $"{invoicePrefix}{(startIndex + i):D5}";

            // Write invoice number and other data to cells
            sheet.Cells[excelRow, 0].PutValue(invoiceNumber);               // Column A
            sheet.Cells[excelRow, 1].PutValue(customers[i].Name);          // Column B
            sheet.Cells[excelRow, 2].PutValue(customers[i].Amount);       // Column C
        }

        // Auto‑fit columns for better readability
        sheet.AutoFitColumns();

        // Save the workbook to a file
        string outputPath = "InvoiceReport.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Invoice report generated and saved to '{outputPath}'.");
    }
}
