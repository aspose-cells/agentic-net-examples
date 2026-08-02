using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

class TableOfContentsWithSmartMarkers
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Add sample worksheets with some data (populated)
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sales";
        sheet1.Cells["A1"].PutValue("Product");
        sheet1.Cells["B1"].PutValue("Quantity");
        sheet1.Cells["A2"].PutValue("Apple");
        sheet1.Cells["B2"].PutValue(150);

        Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
        sheet2.Cells["A1"].PutValue("Item");
        sheet2.Cells["B1"].PutValue("Stock");
        sheet2.Cells["A2"].PutValue("Pen");
        sheet2.Cells["B2"].PutValue(500);

        // Add an empty worksheet (should not appear in TOC)
        Worksheet emptySheet = workbook.Worksheets.Add("EmptySheet");

        // -------------------------------------------------
        // Create a TOC (Table of Contents) worksheet at the beginning
        // -------------------------------------------------
        Worksheet tocSheet = workbook.Worksheets[0];
        tocSheet.Name = "TOC";

        // Header for TOC
        tocSheet.Cells["A1"].PutValue("Table of Contents");
        tocSheet.Cells["A2"].PutValue("Sheet Name");

        // Smart marker that will be repeated for each row in the data source
        // The range starting at A3 will be processed by WorkbookDesigner
        tocSheet.Cells["A3"].PutValue("&=Sheets.Name");

        // -------------------------------------------------
        // Build a DataTable containing names of all populated worksheets (excluding TOC)
        // -------------------------------------------------
        DataTable dtSheets = new DataTable("Sheets");
        dtSheets.Columns.Add("Name", typeof(string));

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];

            // Skip the TOC sheet itself
            if (ws.Name == "TOC")
                continue;

            // Determine if the worksheet has any data
            // MaxDataRow/MaxDataColumn are -1 when the sheet is empty
            if (ws.Cells.MaxDataRow >= 0 && ws.Cells.MaxDataColumn >= 0)
            {
                dtSheets.Rows.Add(ws.Name);
            }
        }

        // -------------------------------------------------
        // Process smart markers using WorkbookDesigner
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dtSheets);
        designer.Process();

        // -------------------------------------------------
        // After processing, add hyperlinks to each sheet name in the TOC
        // -------------------------------------------------
        // Data starts from row index 2 (zero‑based) i.e., Excel row 3
        int startRow = 2;
        for (int i = 0; i < dtSheets.Rows.Count; i++)
        {
            string sheetName = dtSheets.Rows[i]["Name"].ToString();

            // Add a hyperlink that points to cell A1 of the target sheet
            // Hyperlink address format: "#'SheetName'!A1"
            string address = $"#'{sheetName}'!A1";

            // Row index in the worksheet (zero‑based)
            int rowIndex = startRow + i;

            // Column index 0 corresponds to column A
            tocSheet.Hyperlinks.Add(rowIndex, 0, 1, 1, address);
        }

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("Workbook_With_TOC.xlsx");
    }
}