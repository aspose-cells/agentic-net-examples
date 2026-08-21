// Title: C# – Generate Excel Invoice with Line Items, Totals & Company Logo via Aspose.Cells Smart Markers
// Description: Shows how to create an Excel invoice template using Aspose.Cells smart markers, bind an InvoiceHeader (including a logo image) and a list of InvoiceItem objects, repeat rows for each item, calculate the total with a formula, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | smart markers | C# invoice generation | Excel invoice template | logo image insertion | repeating rows | total calculation | WorkbookDesigner | named range | export to PDF
// Common Searches: Aspose.Cells smart markers invoice example C# | insert image into Excel using Aspose.Cells | repeat rows for line items Aspose.Cells | calculate sum column in generated invoice | create invoice template with WorkbookDesigner
// Developer Intent: Generate a formatted Excel invoice by populating a smart‑marker template with header fields, a company logo, and a dynamic collection of line items.
// Use Cases: Automate customer invoicing from order data while preserving brand imagery. | Produce batch invoices for multiple clients using a single reusable template. | Integrate generated invoices into accounting systems with built‑in total calculations. | Export the populated workbook to PDF or other formats for distribution.
// AI Prompts: Add a tax row after the total in the smart‑marker invoice template. | Provide code to convert the generated invoice workbook to PDF with Aspose.Cells. | Explain how to bind additional data sources, such as a customer address, to the same WorkbookDesigner instance.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace InvoiceGenerator
{
    // Data class for invoice header information
    // Shows how to create an Excel invoice template using Aspose.Cells smart markers, bind an InvoiceHeader (including a logo image) and a list of InvoiceItem objects, repeat rows for each item, calculate the total with a formula, and save the workbook as an XLSX file.
    public class InvoiceHeader
    {
        public string? CompanyName { get; set; }
        public string? InvoiceNumber { get; set; }
        public DateTime Date { get; set; }
        public byte[]? Logo { get; set; }   // Image data for the company logo
    }

    // Data class for a single line item
    public class InvoiceItem
    {
        public string? Item { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }

        // Calculated amount (Quantity * Price)
        public decimal Amount => Quantity * Price;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ------------------- Create workbook and template -------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header section with smart markers
                cells["A1"].PutValue("Company:");
                cells["B1"].PutValue("&InvoiceHeader.CompanyName");          // Text marker
                cells["A2"].PutValue("Invoice #:"); 
                cells["B2"].PutValue("&InvoiceHeader.InvoiceNumber");       // Text marker
                cells["A3"].PutValue("Date:");
                cells["B3"].PutValue("&InvoiceHeader.Date");                // Date marker
                cells["A1"].PutValue("&=$Logo");                           // Image marker (logo will be placed here)

                // Column titles for line items
                cells["A5"].PutValue("Item");
                cells["B5"].PutValue("Quantity");
                cells["C5"].PutValue("Unit Price");
                cells["D5"].PutValue("Amount");

                // Row template for line items (will be repeated for each item)
                cells["A6"].PutValue("&Items.Item");
                cells["B6"].PutValue("&Items.Quantity");
                cells["C6"].PutValue("&Items.Price");
                cells["D6"].PutValue("&Items.Amount");

                // Define the range that contains the repeating row.
                // The range must be named "_CellsSmartMarkers" when LineByLine is false.
                Aspose.Cells.Range itemsRange = cells.CreateRange("A6:D6");
                itemsRange.Name = "_CellsSmartMarkers";

                // Total row (after the items)
                cells["C10"].PutValue("Total:");
                cells["D10"].Formula = "SUM(D6:D9)";   // Will sum the generated amount rows

                // ------------------- Prepare data sources -------------------
                // Load logo image (ensure the file exists at the specified path)
                byte[]? logoBytes = null;
                string logoPath = "company_logo.png";
                if (File.Exists(logoPath))
                {
                    logoBytes = File.ReadAllBytes(logoPath);
                }

                var header = new InvoiceHeader
                {
                    CompanyName = "Acme Corp.",
                    InvoiceNumber = "INV-1001",
                    Date = DateTime.Today,
                    Logo = logoBytes
                };

                var items = new List<InvoiceItem>
                {
                    new InvoiceItem { Item = "Widget A", Quantity = 5, Price = 9.99m },
                    new InvoiceItem { Item = "Widget B", Quantity = 3, Price = 14.50m },
                    new InvoiceItem { Item = "Service C", Quantity = 1, Price = 199.00m }
                };

                // ------------------- Configure WorkbookDesigner -------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; range smart markers are used via the named range above.
                };

                // Bind data sources to the smart marker names used in the template
                designer.SetDataSource("InvoiceHeader", header);
                designer.SetDataSource("Items", items);

                // Process the smart markers and populate the workbook
                designer.Process();

                // ------------------- Save the generated invoice -------------------
                string outputPath = "GeneratedInvoice.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Invoice generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error generating invoice: {ex.Message}");
            }
        }
    }
}
