// Title: Create a Table of Contents worksheet with smart markers and internal hyperlinks for each populated sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a TOC sheet that lists every worksheet containing data and inserts a clickable internal link to cell A1 using Aspose.Cells smart markers. | Write C# code to collect worksheet names and addresses into a DataTable, bind it to WorkbookDesigner, and process the smart markers to populate the Table of Contents. | After processing, add Hyperlink objects to the link column so each entry navigates to its sheet, while skipping the TOC worksheet itself.
// Common Searches: asp.net generate excel table of contents with hyperlinks using Aspose.Cells smart markers | c# create dynamic sheet index in Excel workbook excluding TOC sheet | how to bind DataTable to WorkbookDesigner for smart marker processing in Aspose.Cells | add internal hyperlink to Excel cell programmatically with Aspose.Cells .NET | list only populated worksheets in Aspose.Cells and create clickable TOC
// Tags: Aspose.Cells WorkbookDesigner smart markers | C# generate Excel TOC with hyperlinks | populate Table of Contents sheet from DataTable | internal hyperlink address format Excel Aspose | exclude TOC worksheet when building sheet index

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample worksheets, builds a DataTable of non‑empty sheet names and internal addresses, binds it to WorkbookDesigner, uses smart markers to fill a 'Table of Contents' sheet, then adds Hyperlink objects to the link column so each entry points to the corresponding sheet, finally saving the file as TableOfContents.xlsx.
    class TableOfContentsWithSmartMarkers
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // Add sample worksheets with some data (for demonstration purposes)
            // -----------------------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sales";
            sheet1.Cells["A1"].PutValue("Quarter");
            sheet1.Cells["B1"].PutValue("Revenue");
            sheet1.Cells["A2"].PutValue("Q1");
            sheet1.Cells["B2"].PutValue(120000);
            sheet1.Cells["A3"].PutValue("Q2");
            sheet1.Cells["B3"].PutValue(150000);

            Worksheet sheet2 = workbook.Worksheets.Add("Inventory");
            sheet2.Cells["A1"].PutValue("Item");
            sheet2.Cells["B1"].PutValue("Quantity");
            sheet2.Cells["A2"].PutValue("Apples");
            sheet2.Cells["B2"].PutValue(500);
            sheet2.Cells["A3"].PutValue("Oranges");
            sheet2.Cells["B3"].PutValue(300);

            // -----------------------------------------------------------------
            // Add a Table of Contents (TOC) worksheet
            // -----------------------------------------------------------------
            Worksheet tocSheet = workbook.Worksheets.Add("Table of Contents");
            // Place headers
            tocSheet.Cells["A1"].PutValue("Sheet Name");
            tocSheet.Cells["B1"].PutValue("Link");

            // Insert smart markers that will be replaced by the data source
            // The range starting at A2 will be repeated for each row in the data source
            tocSheet.Cells["A2"].PutValue("&=Sheets.Name");
            tocSheet.Cells["B2"].PutValue("&=Sheets.Address");

            // -----------------------------------------------------------------
            // Build a DataTable containing the names of all populated worksheets
            // (excluding the TOC sheet itself)
            // -----------------------------------------------------------------
            DataTable sheetTable = new DataTable("Sheets");
            sheetTable.Columns.Add("Name", typeof(string));
            sheetTable.Columns.Add("Address", typeof(string));

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                // Skip the TOC sheet
                if (ws.Name == tocSheet.Name) continue;

                // Consider a worksheet "populated" if it has at least one non‑empty cell
                if (ws.Cells.MaxDataRow >= 0 && ws.Cells.MaxDataColumn >= 0)
                {
                    // Internal hyperlink address format: "#'SheetName'!A1"
                    string address = $"#'{ws.Name}'!A1";
                    sheetTable.Rows.Add(ws.Name, address);
                }
            }

            // -----------------------------------------------------------------
            // Use WorkbookDesigner to process the smart markers
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource("Sheets", sheetTable);
            designer.Process();

            // -----------------------------------------------------------------
            // After processing, add actual hyperlink objects to the "Link" column
            // -----------------------------------------------------------------
            // Determine how many rows were filled (excluding header)
            int dataRows = sheetTable.Rows.Count;
            for (int row = 2; row < 2 + dataRows; row++)
            {
                // Read the address that was placed by the smart marker
                string linkAddress = tocSheet.Cells[row, 1].StringValue; // column B (index 1)

                // Add hyperlink to the cell in column B
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
                tocSheet.Hyperlinks.Add(row, 1, 1, 1, linkAddress);
            }

            // -----------------------------------------------------------------
            // Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("TableOfContents.xlsx");
        }
    }
}
