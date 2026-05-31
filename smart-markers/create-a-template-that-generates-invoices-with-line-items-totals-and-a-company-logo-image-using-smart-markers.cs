using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace InvoiceGenerator
{
    // Data model for the invoice
    public class InvoiceData
    {
        public string? CompanyName { get; set; }
        public string? CompanyAddress { get; set; }
        public string? InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public byte[]? CompanyLogo { get; set; }          // Image data for the logo
        public List<LineItem>? Items { get; set; }        // Collection of line items
        public decimal SubTotal { get; set; }
        public decimal Tax { get; set; }
        public decimal Total { get; set; }
    }

    // Data model for a single line item
    public class LineItem
    {
        public string? Description { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Amount => Quantity * UnitPrice;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create template workbook --------------------
                var wb = new Workbook();                         // create workbook
                var ws = wb.Worksheets[0];                       // use first worksheet

                // Header area with company information and logo
                ws.Cells["A1"].PutValue("&=Invoice.CompanyLogo");    // logo placeholder (image)
                ws.Cells["B1"].PutValue("&=Invoice.CompanyName");    // company name
                ws.Cells["B2"].PutValue("&=Invoice.CompanyAddress"); // company address
                ws.Cells["E1"].PutValue("Invoice #:");               // static label
                ws.Cells["F1"].PutValue("&=Invoice.InvoiceNumber");  // invoice number
                ws.Cells["E2"].PutValue("Date:");                    // static label
                ws.Cells["F2"].PutValue("&=Invoice.InvoiceDate");    // invoice date

                // Table header for line items
                ws.Cells["A5"].PutValue("Description");
                ws.Cells["B5"].PutValue("Quantity");
                ws.Cells["C5"].PutValue("Unit Price");
                ws.Cells["D5"].PutValue("Amount");

                // Smart markers for line items (repeatable rows)
                ws.Cells["A6"].PutValue("&=Items.Description");
                ws.Cells["B6"].PutValue("&=Items.Quantity");
                ws.Cells["C6"].PutValue("&=Items.UnitPrice");
                ws.Cells["D6"].PutValue("&=Items.Amount");

                // Totals area
                ws.Cells["C10"].PutValue("Subtotal:");
                ws.Cells["D10"].PutValue("&=Invoice.SubTotal");
                ws.Cells["C11"].PutValue("Tax:");
                ws.Cells["D11"].PutValue("&=Invoice.Tax");
                ws.Cells["C12"].PutValue("Total:");
                ws.Cells["D12"].PutValue("&=Invoice.Total");

                // Define the range that contains all smart markers
                // The range must be named "_CellsSmartMarkers" when using range smart markers
                Aspose.Cells.Range smartRange = ws.Cells.CreateRange("A1:D12");
                smartRange.Name = "_CellsSmartMarkers";

                // -------------------- Prepare data source --------------------
                // Load a sample logo image (replace with actual path)
                byte[] logoBytes = File.Exists("logo.png") ? File.ReadAllBytes("logo.png") : Array.Empty<byte>();

                // Sample line items
                var items = new List<LineItem>
                {
                    new LineItem { Description = "Product A", Quantity = 2, UnitPrice = 49.99m },
                    new LineItem { Description = "Product B", Quantity = 1, UnitPrice = 149.50m },
                    new LineItem { Description = "Service C", Quantity = 5, UnitPrice = 30.00m }
                };

                // Calculate totals
                decimal subTotal = 0;
                foreach (var it in items) subTotal += it.Amount;
                decimal tax = subTotal * 0.10m; // 10% tax
                decimal total = subTotal + tax;

                // Invoice data object
                var invoice = new InvoiceData
                {
                    CompanyName = "Acme Corp.",
                    CompanyAddress = "123 Business Rd., Metropolis",
                    InvoiceNumber = "INV-2023-001",
                    InvoiceDate = DateTime.Today,
                    CompanyLogo = logoBytes,
                    Items = items,
                    SubTotal = subTotal,
                    Tax = tax,
                    Total = total
                };

                // -------------------- Process smart markers --------------------
                var designer = new WorkbookDesigner
                {
                    Workbook = wb
                };

                // Bind data sources
                designer.SetDataSource("Invoice", invoice);
                designer.SetDataSource("Items", invoice.Items);

                // Process the template (true = preserve unrecognized markers, not needed here)
                designer.Process(true);

                // -------------------- Save the generated invoice --------------------
                string outputPath = "GeneratedInvoice.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Invoice generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error generating invoice: {ex.Message}");
            }
        }
    }
}