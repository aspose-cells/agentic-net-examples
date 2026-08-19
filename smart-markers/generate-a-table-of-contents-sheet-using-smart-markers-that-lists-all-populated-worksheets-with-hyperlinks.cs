// Title: Create a TOC worksheet with smart markers and hyperlinks using Aspose.Cells for .NET
// Description: This example builds an Excel workbook, adds sample data sheets, inserts a Table of Contents sheet at the first position, and uses Aspose.Cells smart markers to automatically list every populated worksheet with a clickable link to its A1 cell. The WorkbookDesigner processes the _CellsSmartMarkers range, generating a dynamic navigation index that excludes the TOC sheet itself.
// Keywords: Aspose.Cells | smart markers | Table of Contents | Excel TOC | C# .NET | WorkbookDesigner | hyperlink formula | populate sheet list | dynamic index
// Common Searches: Aspose.Cells smart markers TOC example | C# generate Excel table of contents with hyperlinks | list only non‑empty worksheets using Aspose.Cells | WorkbookDesigner create dynamic index sheet | how to add hyperlink to sheet in Aspose.Cells
// Developer Intent: Automatically generate a TOC sheet that lists each populated worksheet and provides a clickable link to its first cell.
// Use Cases: Create a navigation index for workbooks containing many data sheets. | Produce a dynamic report where only sheets with content appear in the TOC. | Offer end‑users quick access to populated worksheets in generated Excel files.
// AI Prompts: Write C# code with Aspose.Cells to add a TOC worksheet that lists all non‑empty sheets and creates hyperlinks using smart markers. | Explain the role of the _CellsSmartMarkers range when processing a table of contents with WorkbookDesigner. | Suggest how to add sheet order numbers and custom link text to the generated TOC.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsTOCExample
{
    // This example builds an Excel workbook, adds sample data sheets, inserts a Table of Contents sheet at the first position, and uses Aspose.Cells smart markers to automatically list every populated worksheet with a clickable link to its A1 cell. The WorkbookDesigner processes the _CellsSmartMarkers range, generating a dynamic navigation index that excludes the TOC sheet itself.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample worksheets with some data
                for (int i = 1; i <= 3; i++)
                {
                    // Add a new worksheet with the specified name
                    Worksheet ws = workbook.Worksheets.Add($"Sheet{i}");
                    // Populate a cell to make the sheet "populated"
                    ws.Cells["A1"].PutValue($"Data in {ws.Name}");
                }

                // Insert a Table of Contents (TOC) sheet at the first position
                Worksheet tocSheet = workbook.Worksheets[0];
                tocSheet.Name = "TOC";

                // Header for the TOC sheet
                tocSheet.Cells["A1"].PutValue("Table of Contents");
                tocSheet.Cells["A2"].PutValue("Sheet Name");
                tocSheet.Cells["B2"].PutValue("Link");

                // Smart marker rows – these will be expanded by WorkbookDesigner
                // Column A: sheet name smart marker
                tocSheet.Cells["A3"].PutValue("&=[Sheets].SheetName");
                // Column B: hyperlink formula that points to cell A1 of the target sheet
                tocSheet.Cells["B3"].PutValue("=HYPERLINK(\"#'\" & [Sheets].SheetName & \"'!A1\",\"Go\")");

                // Define the smart marker range (required name: _CellsSmartMarkers)
                tocSheet.Cells.CreateRange("A3:B3").Name = "_CellsSmartMarkers";

                // Build a DataTable that contains the names of all populated worksheets (excluding the TOC sheet)
                DataTable sheetTable = new DataTable("Sheets");
                sheetTable.Columns.Add("SheetName", typeof(string));

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the TOC sheet itself
                    if (ws.Name == "TOC")
                        continue;

                    // Consider a worksheet "populated" if it has at least one non‑empty cell
                    if (ws.Cells.MaxDataRow >= 0 && ws.Cells.MaxDataColumn >= 0)
                    {
                        sheetTable.Rows.Add(ws.Name);
                    }
                }

                // Use WorkbookDesigner to process the smart markers with the prepared data source
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(sheetTable);
                designer.Process();

                // Save the final workbook
                string outputPath = "TOC_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
