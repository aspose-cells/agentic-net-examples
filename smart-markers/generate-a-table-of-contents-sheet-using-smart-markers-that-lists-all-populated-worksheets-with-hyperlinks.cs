// Title: Create a Table of Contents worksheet with hyperlinks to populated sheets using Aspose.Cells smart markers in C#
// Common Searches: C# Aspose.Cells smart markers to build a worksheet index with clickable links | how to list only non‑empty worksheets in an Excel file using Aspose.Cells | generate a Table of Contents sheet that links to each sheet's A1 cell with Aspose.Cells | using WorkbookDesigner and DataTable to populate smart markers for a TOC in Aspose.Cells | exclude empty worksheets when creating an Excel TOC with Aspose.Cells .NET
// Tags: smart markers generate worksheet index | Aspose.Cells hyperlink formula to sheet cell | exclude empty worksheets Aspose.Cells | WorkbookDesigner set DataTable source | C# create Excel table of contents with links

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, adds populated worksheets, and skips an empty sheet. It then adds a "Table of Contents" sheet, inserts smart markers for sheet names and hyperlink formulas, builds a DataTable containing only non‑empty worksheets with HYPERLINK formulas that point to each sheet's A1 cell, applies the data via WorkbookDesigner, and saves the file as TableOfContents.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Add sample populated worksheets (for demonstration)
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
        sheet1.Name = "Sales";
        sheet1.Cells["A1"].PutValue("Product");
        sheet1.Cells["B1"].PutValue("Quantity");
        sheet1.Cells["A2"].PutValue("Apple");
        sheet1.Cells["B2"].PutValue(120);

        Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
        sheet2.Name = "Inventory";
        sheet2.Cells["A1"].PutValue("Item");
        sheet2.Cells["B1"].PutValue("Stock");
        sheet2.Cells["A2"].PutValue("Banana");
        sheet2.Cells["B2"].PutValue(80);

        // Add an empty worksheet that should be ignored
        Worksheet emptySheet = workbook.Worksheets[workbook.Worksheets.Add()];
        emptySheet.Name = "EmptySheet";

        // -------------------------------------------------
        // Create the Table of Contents sheet
        // -------------------------------------------------
        Worksheet tocSheet = workbook.Worksheets[workbook.Worksheets.Add()];
        tocSheet.Name = "Table of Contents";

        // Set header titles
        tocSheet.Cells["A1"].PutValue("Worksheet");
        tocSheet.Cells["B1"].PutValue("Link");

        // Insert smart markers starting from row 2
        // &="SheetName" will be replaced with the actual sheet name
        // &="Link" will be replaced with the hyperlink formula
        tocSheet.Cells["A2"].PutValue("&=SheetName");
        tocSheet.Cells["B2"].PutValue("&=Link");

        // -------------------------------------------------
        // Prepare data source for the smart markers
        // -------------------------------------------------
        DataTable dt = new DataTable();
        dt.Columns.Add("SheetName", typeof(string));
        dt.Columns.Add("Link", typeof(string));

        // Iterate through all worksheets except the TOC sheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Skip the TOC sheet itself
            if (ws.Name == tocSheet.Name)
                continue;

            // Determine if the worksheet is populated (has at least one non‑empty cell)
            bool isPopulated = ws.Cells.MaxDataRow >= 0 && ws.Cells.MaxDataColumn >= 0;
            if (!isPopulated)
                continue; // ignore empty sheets

            // Build the hyperlink formula that points to cell A1 of the target sheet
            // The formula uses the HYPERLINK function with a sheet reference
            string hyperlinkFormula = $"=HYPERLINK(\"#'{ws.Name}'!A1\",\"{ws.Name}\")";

            // Add a row to the DataTable
            dt.Rows.Add(ws.Name, hyperlinkFormula);
        }

        // -------------------------------------------------
        // Apply the smart markers using WorkbookDesigner
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // replaces smart markers with actual data

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("TableOfContents.xlsx");
    }
}
