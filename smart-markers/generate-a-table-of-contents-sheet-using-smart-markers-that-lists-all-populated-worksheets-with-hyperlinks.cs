// Title: Create an Excel Table of Contents sheet with smart markers and hyperlinks to each populated worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to generate a table of contents worksheet with hyperlinks using Aspose.Cells in C# | Aspose.Cells smart markers example for listing populated sheets | C# code to add a TOC sheet that links to each worksheet in Excel | using WorkbookDesigner to create a dynamic table of contents in an Excel file | filter non‑empty worksheets and create hyperlinks with Aspose.Cells
// Tags: Aspose.Cells smart markers table of contents | C# WorkbookDesigner generate worksheet list | Excel hyperlink formula to sheet A1 via smart marker | populate TOC from non‑empty worksheets Aspose.Cells | create TOC worksheet using Aspose.Cells .NET

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example builds an Excel workbook, adds sample worksheets, and inserts a 'Table of Contents' sheet. Smart markers (&=SheetName and a HYPERLINK formula) are placed in the TOC sheet. A DataTable collects the names of all worksheets that contain at least one non‑null cell (excluding the TOC itself). The DataTable is supplied to a WorkbookDesigner, which processes the smart markers to output each sheet name and a clickable link to its A1 cell. The final workbook is saved as TableOfContents.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------------------------------------------------
                // Add sample populated worksheets (for demonstration)
                // In real scenario these worksheets already exist
                // ---------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet1.Name = "Sales";
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Amount");
                sheet1.Cells["A2"].PutValue("Widget");
                sheet1.Cells["B2"].PutValue(12345);

                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Name = "Inventory";
                sheet2.Cells["A1"].PutValue("Item");
                sheet2.Cells["B1"].PutValue("Quantity");
                sheet2.Cells["A2"].PutValue("Gadget");
                sheet2.Cells["B2"].PutValue(250);

                // ---------------------------------------------------
                // Add the Table of Contents (TOC) worksheet
                // ---------------------------------------------------
                int tocIndex = workbook.Worksheets.Add();
                Worksheet tocSheet = workbook.Worksheets[tocIndex];
                tocSheet.Name = "Table of Contents";

                // Header
                tocSheet.Cells["A1"].PutValue("Table of Contents");

                // Smart marker header row
                tocSheet.Cells["A2"].PutValue("Sheet Name");
                tocSheet.Cells["B2"].PutValue("Link");

                // Smart marker data row (starting at row 3)
                // Column A will receive the sheet name
                tocSheet.Cells["A3"].PutValue("&=SheetName");

                // Column B will receive a hyperlink formula that points to the sheet's A1 cell
                // The formula uses the sheet name inserted by the smart marker
                tocSheet.Cells["B3"].PutValue("&=HYPERLINK(\"#'\" & SheetName & \"'!A1\", SheetName)");

                // ---------------------------------------------------
                // Prepare data source for the smart markers
                // ---------------------------------------------------
                DataTable dt = new DataTable();
                dt.Columns.Add("SheetName", typeof(string));

                // Collect all worksheets that contain at least one non‑empty cell (populated)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the TOC sheet itself
                    if (ws.Name == tocSheet.Name)
                        continue;

                    bool hasData = false;
                    foreach (Cell cell in ws.Cells)
                    {
                        if (cell.Type != CellValueType.IsNull)
                        {
                            hasData = true;
                            break;
                        }
                    }

                    if (hasData)
                    {
                        DataRow row = dt.NewRow();
                        row["SheetName"] = ws.Name;
                        dt.Rows.Add(row);
                    }
                }

                // ---------------------------------------------------
                // Process the smart markers using WorkbookDesigner
                // ---------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                designer.SetDataSource(ds);   // Set the data source
                designer.Process();           // Apply smart markers

                // ---------------------------------------------------
                // Save the workbook
                // ---------------------------------------------------
                string outputPath = "TableOfContents.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
