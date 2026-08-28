// Title: Create an Excel invoice with dynamic line items, total calculation, and embedded logo using Aspose.Cells smart markers (C#)
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to populate an invoice template with company name, invoice number, date, and a logo image via smart markers. | Show how to define a range smart marker in Aspose.Cells that repeats rows for each invoice item and automatically fills description, quantity, unit price, and amount. | Demonstrate calculating the total amount from a list of invoice items in C# and inserting it into the Excel file using a smart marker.
// Common Searches: aspnet c# embed logo image in Excel invoice using Aspose.Cells smart markers | aspose.cells repeat rows for a collection with range smart markers c# | generate invoice spreadsheet with totals and line items using Aspose.Cells WorkbookDesigner | c# smart markers image placeholder in Excel template | calculate invoice total in Aspose.Cells template c#
// Tags: WorkbookDesigner image smart marker | dynamic row replication using smart markers | bind object collection to Excel template C# | invoice total field smart marker | logo image byte array insertion

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace InvoiceGenerator
{
    // Data model for invoice
    // The example builds an Excel workbook, inserts smart markers for static fields (company name, invoice number, date), places an image smart marker for the logo, defines a repeatable range smart marker for line items, binds an InvoiceData object and its Items list to WorkbookDesigner, processes the markers, calculates the total amount, and saves the populated invoice as GeneratedInvoice.xlsx.
    public class InvoiceData
    {
        public string CompanyName { get; set; }
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public byte[] Logo { get; set; }               // Image data for company logo
        public List<InvoiceItem> Items { get; set; }   // Collection of line items
        public decimal TotalAmount { get; set; }       // Calculated total
    }

    // Data model for a line item
    public class InvoiceItem
    {
        public string Description { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Amount => Quantity * UnitPrice;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Place smart markers for static fields
            sheet.Cells["A1"].PutValue("Company:");
            sheet.Cells["B1"].PutValue("&=CompanyName");          // Company name
            sheet.Cells["A2"].PutValue("Invoice #:");           
            sheet.Cells["B2"].PutValue("&=InvoiceNumber");        // Invoice number
            sheet.Cells["A3"].PutValue("Date:");
            sheet.Cells["B3"].PutValue("&=InvoiceDate");          // Invoice date

            // 3. Placeholder for company logo (image smart marker)
            sheet.Cells["A5"].PutValue("&=$Logo");                // Image marker

            // 4. Header for line items table
            sheet.Cells["A7"].PutValue("Description");
            sheet.Cells["B7"].PutValue("Quantity");
            sheet.Cells["C7"].PutValue("Unit Price");
            sheet.Cells["D7"].PutValue("Amount");

            // 5. Define a range that will be repeated for each line item
            //    The range includes the row with smart markers for the item fields
            //    Name the range "_CellsSmartMarkers" to enable range smart markers
            sheet.Cells.CreateRange("A8:D18").Name = "_CellsSmartMarkers";

            // 6. Set smart markers inside the range (they will be repeated per item)
            sheet.Cells["A8"].PutValue("&=$Items.Description");
            sheet.Cells["B8"].PutValue("&=$Items.Quantity");
            sheet.Cells["C8"].PutValue("&=$Items.UnitPrice");
            sheet.Cells["D8"].PutValue("&=$Items.Amount");

            // 7. Place total amount label and smart marker
            sheet.Cells["C20"].PutValue("Total:");
            sheet.Cells["D20"].PutValue("&=TotalAmount");

            // 8. Prepare sample data
            InvoiceData invoice = new InvoiceData
            {
                CompanyName = "Acme Corp.",
                InvoiceNumber = "INV-1001",
                InvoiceDate = DateTime.Today,
                // Load logo image bytes (ensure the file exists at the specified path)
                Logo = File.Exists("logo.png") ? File.ReadAllBytes("logo.png") : new byte[0],
                Items = new List<InvoiceItem>
                {
                    new InvoiceItem { Description = "Widget A", Quantity = 5, UnitPrice = 9.99m },
                    new InvoiceItem { Description = "Widget B", Quantity = 3, UnitPrice = 14.50m },
                    new InvoiceItem { Description = "Service C", Quantity = 1, UnitPrice = 199.00m }
                }
            };
            // Calculate total amount
            decimal total = 0;
            foreach (var item in invoice.Items) total += item.Amount;
            invoice.TotalAmount = total;

            // 9. Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook)
            {
                // Use range smart markers (LineByLine must be false)
                LineByLine = false
            };

            // 10. Bind data sources
            designer.SetDataSource("Invoice", invoice);          // For static fields and logo
            designer.SetDataSource("Items", invoice.Items);      // For line items

            // 11. Process smart markers (true = preserve unrecognized markers)
            designer.Process(true);

            // 12. Save the generated invoice
            workbook.Save("GeneratedInvoice.xlsx");
        }
    }
}
