// Title: Generate an Excel invoice with line items, totals, and a company logo using Aspose.Cells smart markers in C#
// Common Searches: c# create invoice spreadsheet with Aspose.Cells smart markers | insert company logo image into Excel using Aspose.Cells image smart marker | populate invoice line items collection with WorkbookDesigner in C# | calculate subtotal tax and total in Aspose.Cells generated invoice | auto fit columns after processing smart markers in Aspose.Cells
// Tags: smart markers invoice generation Aspose.Cells | embed logo via smart marker | WorkbookDesigner data source binding | invoice totals calculation Aspose.Cells | column width optimization Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace InvoiceGenerator
{
    // Data models for the smart markers
    // The example builds an Excel invoice template by placing smart marker placeholders for a company logo, company details, line‑item fields, and total values, then saves it. An Invoice object is populated with company information, a list of InvoiceItem objects, optional logo bytes, and calculated SubTotal, Tax, and Total. WorkbookDesigner loads the template, binds the Invoice object as a data source, processes the smart markers to fill the sheet, and finally saves the completed invoice workbook.
    public class Invoice
    {
        public string CompanyName { get; set; } = string.Empty;
        public string InvoiceNumber { get; set; } = string.Empty;
        public DateTime InvoiceDate { get; set; }
        public byte[]? CompanyLogo { get; set; }          // Image data (nullable if file missing)
        public List<InvoiceItem> Items { get; set; } = new List<InvoiceItem>();
        public decimal SubTotal { get; set; }
        public decimal Tax { get; set; }
        public decimal Total { get; set; }
    }

    public class InvoiceItem
    {
        public string Description { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal LineTotal => Quantity * UnitPrice;
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // STEP 1: Create the template workbook with smart markers
                // -----------------------------------------------------------------
                Workbook templateWb = new Workbook();
                Worksheet ws = templateWb.Worksheets[0];
                Cells cells = ws.Cells;

                // Company logo placeholder (image smart marker)
                cells["A1"].PutValue("&<#Image:CompanyLogo#>"); // Image will be inserted here

                // Company details
                cells["B1"].PutValue("&<#CompanyName#>");
                cells["B2"].PutValue("Invoice #: &<#InvoiceNumber#>");
                cells["B3"].PutValue("Date: &<#InvoiceDate#>");

                // Header for line items table
                cells["A5"].PutValue("Description");
                cells["B5"].PutValue("Quantity");
                cells["C5"].PutValue("Unit Price");
                cells["D5"].PutValue("Line Total");

                // Smart markers for the collection "Items"
                cells["A6"].PutValue("&<#Items.Description#>");
                cells["B6"].PutValue("&<#Items.Quantity#>");
                cells["C6"].PutValue("&<#Items.UnitPrice#>");
                cells["D6"].PutValue("&<#Items.LineTotal#>");

                // Totals (simple placeholders, will be filled from the data source)
                cells["C8"].PutValue("Subtotal:");
                cells["D8"].PutValue("&<#SubTotal#>");

                cells["C9"].PutValue("Tax:");
                cells["D9"].PutValue("&<#Tax#>");

                cells["C10"].PutValue("Total:");
                cells["D10"].PutValue("&<#Total#>");

                // Adjust column widths for better appearance
                ws.AutoFitColumns();

                // Save the template
                string templatePath = "InvoiceTemplate.xlsx";
                templateWb.Save(templatePath, SaveFormat.Xlsx);

                // -----------------------------------------------------------------
                // STEP 2: Prepare the data source for the smart markers
                // -----------------------------------------------------------------
                Invoice invoice = new Invoice
                {
                    CompanyName = "Acme Corp.",
                    InvoiceNumber = "INV-1001",
                    InvoiceDate = DateTime.Today,
                    Items = new List<InvoiceItem>
                    {
                        new InvoiceItem { Description = "Widget A", Quantity = 5, UnitPrice = 9.99m },
                        new InvoiceItem { Description = "Widget B", Quantity = 3, UnitPrice = 19.95m },
                        new InvoiceItem { Description = "Service C", Quantity = 1, UnitPrice = 149.00m }
                    }
                };

                // Load logo if the file exists
                string logoPath = "CompanyLogo.png";
                if (File.Exists(logoPath))
                {
                    invoice.CompanyLogo = File.ReadAllBytes(logoPath);
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Image placeholder will be empty.");
                }

                // Calculate totals
                invoice.SubTotal = 0;
                foreach (var item in invoice.Items)
                {
                    invoice.SubTotal += item.LineTotal;
                }
                invoice.Tax = Math.Round(invoice.SubTotal * 0.15m, 2); // 15% tax
                invoice.Total = invoice.SubTotal + invoice.Tax;

                // -----------------------------------------------------------------
                // STEP 3: Process the template with WorkbookDesigner (smart markers)
                // -----------------------------------------------------------------
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file '{templatePath}' was not created.");
                }

                Workbook resultWb = new Workbook(templatePath);
                WorkbookDesigner designer = new WorkbookDesigner(resultWb);
                designer.SetDataSource("Invoice", invoice);
                designer.Process();

                // -----------------------------------------------------------------
                // STEP 4: Save the final invoice
                // -----------------------------------------------------------------
                string outputPath = "Invoice_1001.xlsx";
                resultWb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Invoice generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
