// Title: Generate an Excel invoice template with line items, totals, and a company logo using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that creates a new workbook, merges cells for a logo placeholder, and inserts the smart‑marker image tag (&="CompanyLogo") with Aspose.Cells. | Add smart‑marker fields for company name, address, phone, invoice number, date, and a repeatable Items collection (ItemName, Description, Quantity, UnitPrice, LineTotal) to the worksheet. | Style the header row in bold, set appropriate column widths, place subtotal, tax, and total smart‑marker tags, and save the workbook as InvoiceTemplate.xlsx.
// Common Searches: how to insert a company logo using Aspose.Cells smart markers in a C# Excel invoice | c# Aspose.Cells repeatable smart marker block for invoice line items collection | setting column widths and merging cells for logo placeholder with Aspose.Cells smart markers | calculating subtotal tax total in an Excel invoice template using Aspose.Cells smart markers | Aspose.Cells smart marker syntax for image and text placeholders in invoice templates
// Tags: Aspose.Cells smart‑marker image placeholder | C# generate Excel invoice template | repeatable collection smart markers Aspose.Cells | merge cells for logo Aspose.Cells | apply header style Aspose.Cells | subtotal tax total smart markers Excel

using Aspose.Cells;
using System;

// The example creates a new workbook, configures column widths, merges cells for a logo placeholder using the smart‑marker syntax &="CompanyLogo", inserts smart markers for company details, invoice number, and date, defines a header row, adds a repeatable block referencing an Items collection for line items, places subtotal, tax, and total smart markers, applies bold styling to the header, and saves the file as InvoiceTemplate.xlsx.
class InvoiceTemplateGenerator
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set column widths for a clean layout
        sheet.Cells.SetColumnWidth(0, 20); // Item
        sheet.Cells.SetColumnWidth(1, 30); // Description
        sheet.Cells.SetColumnWidth(2, 15); // Quantity
        sheet.Cells.SetColumnWidth(3, 15); // Unit Price
        sheet.Cells.SetColumnWidth(4, 20); // Line Total

        // -------------------------------------------------
        // Company logo (smart marker for image)
        // The placeholder will be replaced with an image when the template is processed
        // Smart marker syntax for image: &="CompanyLogo"
        // Merge a few cells to give the logo some space
        // -------------------------------------------------
        sheet.Cells.Merge(0, 0, 4, 2);
        sheet.Cells["A1"].PutValue("&=\"CompanyLogo\"");

        // -------------------------------------------------
        // Company information (text smart markers)
        // -------------------------------------------------
        sheet.Cells["C1"].PutValue("&=\"CompanyName\"");
        sheet.Cells["C2"].PutValue("&=\"CompanyAddress\"");
        sheet.Cells["C3"].PutValue("&=\"CompanyPhone\"");

        // -------------------------------------------------
        // Invoice header
        // -------------------------------------------------
        sheet.Cells["A5"].PutValue("Invoice #:");               // static label
        sheet.Cells["B5"].PutValue("&=\"InvoiceNumber\"");     // smart marker
        sheet.Cells["A6"].PutValue("Date:");                    // static label
        sheet.Cells["B6"].PutValue("&=\"InvoiceDate\"");       // smart marker

        // -------------------------------------------------
        // Table header for line items
        // -------------------------------------------------
        int headerRow = 8;
        sheet.Cells[headerRow, 0].PutValue("Item");
        sheet.Cells[headerRow, 1].PutValue("Description");
        sheet.Cells[headerRow, 2].PutValue("Quantity");
        sheet.Cells[headerRow, 3].PutValue("Unit Price");
        sheet.Cells[headerRow, 4].PutValue("Line Total");

        // Apply bold style to the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        for (int col = 0; col <= 4; col++)
        {
            sheet.Cells[headerRow, col].SetStyle(headerStyle);
        }

        // -------------------------------------------------
        // Smart markers for line items (repeatable block)
        // Use a collection named "Items" in the data source.
        // Each field is referenced as Items.FieldName
        // -------------------------------------------------
        int itemRow = headerRow + 1;
        sheet.Cells[itemRow, 0].PutValue("&=\"Items.ItemName\"");
        sheet.Cells[itemRow, 1].PutValue("&=\"Items.Description\"");
        sheet.Cells[itemRow, 2].PutValue("&=\"Items.Quantity\"");
        sheet.Cells[itemRow, 3].PutValue("&=\"Items.UnitPrice\"");
        sheet.Cells[itemRow, 4].PutValue("&=\"Items.LineTotal\"");

        // -------------------------------------------------
        // Totals section (simple smart markers)
        // -------------------------------------------------
        int totalsStartRow = itemRow + 2; // leave one empty row after items
        sheet.Cells[totalsStartRow, 3].PutValue("Subtotal:");
        sheet.Cells[totalsStartRow, 4].PutValue("&=\"Subtotal\"");
        sheet.Cells[totalsStartRow + 1, 3].PutValue("Tax:");
        sheet.Cells[totalsStartRow + 1, 4].PutValue("&=\"Tax\"");
        sheet.Cells[totalsStartRow + 2, 3].PutValue("Total:");
        sheet.Cells[totalsStartRow + 2, 4].PutValue("&=\"Total\"");

        // -------------------------------------------------
        // Save the template workbook
        // -------------------------------------------------
        workbook.Save("InvoiceTemplate.xlsx");
    }
}
