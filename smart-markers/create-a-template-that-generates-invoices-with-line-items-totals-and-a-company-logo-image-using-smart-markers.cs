// Title: Create Excel Invoice with Logo, Line Items & Totals Using Aspose.Cells Smart Markers (C#)
// Description: Demonstrates how to build an Excel invoice template, insert a logo image, add header fields, generate a repeating line‑item table, calculate subtotal, tax and grand total, and populate the workbook with an InvoiceData object using Aspose.Cells smart markers and range smart markers.
// Keywords: Aspose.Cells smart markers | C# invoice generation | Excel invoice template | populate image with smart markers | range smart markers example | dynamic line items Excel | calculate tax Aspose.Cells | export invoice to XLSX
// Common Searches: Aspose.Cells create invoice with logo | C# smart markers repeat rows | Excel invoice template using Aspose.Cells | bind collection to smart markers C# | calculate totals in Aspose.Cells invoice
// Developer Intent: Generate a populated Excel invoice by defining smart markers for a logo, header data, a repeatable line‑item section, and total rows, then binding an InvoiceData object and processing the template.
// Use Cases: Produce printable invoices for multiple customers, each with its own logo and item list. | Automate monthly billing statements with automatic tax calculation and currency formatting. | Batch‑process invoices from a database, creating a separate XLSX file for each record.
// AI Prompts: Write C# code that reads a logo file, creates a range smart marker invoice template, binds an InvoiceData object, and saves the result as a PDF using Aspose.Cells. | Update the sample to format Unit Price, Total, Subtotal, Tax, and Grand Total cells as currency with two decimal places and apply bold styling to the header row. | Add error handling that substitutes a default placeholder image when the logo file is missing and logs a warning before generating the invoice.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace InvoiceGenerator
{
    // Data model for a line item
    // Demonstrates how to build an Excel invoice template, insert a logo image, add header fields, generate a repeating line‑item table, calculate subtotal, tax and grand total, and populate the workbook with an InvoiceData object using Aspose.Cells smart markers and range smart markers.
    public class LineItem
    {
        public string Product { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public double UnitPrice { get; set; }
        public double Total => Quantity * UnitPrice;
    }

    // Data model for the invoice
    public class InvoiceData
    {
        public string InvoiceNumber { get; set; } = string.Empty;
        public DateTime InvoiceDate { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public byte[] Logo { get; set; } = Array.Empty<byte>();
        public List<LineItem> LineItems { get; set; } = new List<LineItem>();
        public double Subtotal { get; set; }
        public double Tax { get; set; }
        public double GrandTotal { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook (template) ----------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // ---------- Place smart markers ----------
                // Company logo placeholder (image)
                cells["A1"].PutValue("&=$Logo");

                // Invoice header fields
                cells["A3"].PutValue("Invoice #:");          // static label
                cells["B3"].PutValue("&=$InvoiceNumber");    // smart marker
                cells["A4"].PutValue("Date:");               // static label
                cells["B4"].PutValue("&=$InvoiceDate");      // smart marker
                cells["A5"].PutValue("Customer:");           // static label
                cells["B5"].PutValue("&=$CustomerName");     // smart marker

                // Line items table header
                cells["A7"].PutValue("Product");
                cells["B7"].PutValue("Quantity");
                cells["C7"].PutValue("Unit Price");
                cells["D7"].PutValue("Total");

                // Line items smart markers (will be repeated for each item)
                cells["A8"].PutValue("&=$LineItems.Product");
                cells["B8"].PutValue("&=$LineItems.Quantity");
                cells["C8"].PutValue("&=$LineItems.UnitPrice");
                cells["D8"].PutValue("&=$LineItems.Total");

                // Totals section
                cells["C10"].PutValue("Subtotal:");
                cells["D10"].PutValue("&=$Subtotal");
                cells["C11"].PutValue("Tax:");
                cells["D11"].PutValue("&=$Tax");
                cells["C12"].PutValue("Grand Total:");
                cells["D12"].PutValue("&=$GrandTotal");

                // Define the range that contains all smart markers and name it "_CellsSmartMarkers"
                // This is required when using range smart markers.
                Aspose.Cells.Range smartRange = cells.CreateRange("A1:D12");
                smartRange.Name = "_CellsSmartMarkers";

                // ---------- Prepare sample data ----------
                // Load a logo image file into a byte array (replace with actual path if needed)
                byte[] logoBytes = Array.Empty<byte>();
                const string logoPath = "logo.png";
                if (File.Exists(logoPath))
                {
                    try
                    {
                        logoBytes = File.ReadAllBytes(logoPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to read logo file: {ex.Message}");
                    }
                }

                var invoice = new InvoiceData
                {
                    InvoiceNumber = "INV-1001",
                    InvoiceDate = DateTime.Today,
                    CustomerName = "Acme Corporation",
                    Logo = logoBytes,
                    LineItems = new List<LineItem>
                    {
                        new LineItem { Product = "Widget A", Quantity = 5, UnitPrice = 9.99 },
                        new LineItem { Product = "Widget B", Quantity = 3, UnitPrice = 14.50 },
                        new LineItem { Product = "Service C", Quantity = 1, UnitPrice = 199.00 }
                    }
                };

                // Calculate totals
                invoice.Subtotal = 0;
                foreach (var item in invoice.LineItems)
                {
                    invoice.Subtotal += item.Total;
                }
                invoice.Tax = Math.Round(invoice.Subtotal * 0.07, 2); // 7% tax
                invoice.GrandTotal = invoice.Subtotal + invoice.Tax;

                // ---------- Set up WorkbookDesigner and bind data ----------
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Use range smart markers (LineByLine is obsolete, but kept for compatibility)
                designer.LineByLine = false;

                // Bind the main invoice object and the collection of line items
                designer.SetDataSource("Invoice", invoice);
                designer.SetDataSource("LineItems", invoice.LineItems);

                // Process the smart markers
                designer.Process();

                // ---------- Save the generated invoice ----------
                const string outputPath = "GeneratedInvoice.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Invoice generated successfully: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
