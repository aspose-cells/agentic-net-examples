// Title: Process smart markers in a hidden worksheet with WorkbookDesigner and then unhide the sheet using Aspose.Cells for .NET
// Common Searches: how to process smart markers on a hidden worksheet using Aspose.Cells WorkbookDesigner in C# | unhide a hidden sheet after WorkbookDesigner.Process execution with Aspose.Cells | enable UpdateReference to refresh formulas that refer to hidden worksheets in Aspose.Cells | populate smart markers from a DataTable for hidden sheet processing in Aspose.Cells | preserve unknown smart markers while processing hidden worksheets with Aspose.Cells
// Tags: WorkbookDesigner hidden worksheet smart marker handling | make hidden sheets visible after Aspose.Cells processing | Aspose.Cells UpdateReference for hidden sheet formulas | C# bind DataTable to smart markers in Aspose.Cells | save processed workbook as XLSX with Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace WorkbookDesignerHiddenSmartMarkers
{
    // The example creates a workbook, adds a hidden worksheet containing smart markers, fills a DataTable that matches those markers, assigns the workbook and data source to a WorkbookDesigner, enables reference updates, processes the smart markers while preserving unrecognized ones, then makes all worksheets visible before saving the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add a hidden worksheet that contains
                //    smart markers.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Add a new worksheet that will be hidden
                int hiddenSheetIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
                hiddenSheet.Name = "HiddenData";

                // Hide the worksheet
                hiddenSheet.IsVisible = false;   // false = hidden, true = visible

                // Put some smart markers into the hidden sheet.
                // The syntax "&=Data.ColumnName" tells Aspose.Cells to replace the
                // marker with the value from the data source.
                hiddenSheet.Cells["A1"].PutValue("&=Data.Name");
                hiddenSheet.Cells["B1"].PutValue("&=Data.Age");
                hiddenSheet.Cells["A2"].PutValue("&=Data.Name");
                hiddenSheet.Cells["B2"].PutValue("&=Data.Age");

                // -----------------------------------------------------------------
                // 2. Prepare a data source (DataTable) that matches the smart markers.
                // -----------------------------------------------------------------
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));

                dt.Rows.Add("Alice", 30);
                dt.Rows.Add("Bob", 45);
                dt.Rows.Add("Charlie", 28);

                // -----------------------------------------------------------------
                // 3. Configure WorkbookDesigner, assign the workbook and the data source.
                // -----------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = workbook;
                designer.SetDataSource(dt);               // DataTable name defaults to its TableName ("Data")
                designer.UpdateReference = true;          // Ensure formulas referencing hidden sheet are updated

                // -----------------------------------------------------------------
                // 4. Process smart markers. The boolean parameter indicates whether
                //    unrecognized markers should be preserved (true = preserve).
                // -----------------------------------------------------------------
                designer.Process(true);

                // -----------------------------------------------------------------
                // 5. After processing, unhide all worksheets so the user can see the results.
                // -----------------------------------------------------------------
                foreach (Worksheet ws in designer.Workbook.Worksheets)
                {
                    ws.IsVisible = true;
                }

                // -----------------------------------------------------------------
                // 6. Save the processed workbook.
                // -----------------------------------------------------------------
                string outputPath = "ProcessedWorkbook.xlsx";
                designer.Workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook processed successfully. Saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
