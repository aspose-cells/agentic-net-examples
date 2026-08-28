// Title: Transpose a column‑wise data range and process a named smart‑marker range with WorkbookDesigner in Aspose.Cells for .NET
// AI Prompts: Transpose the raw range E1:F3, export it to a DataTable, assign it to the "Data" source, and process only the smart‑marker range A2:C2 using WorkbookDesigner. | Create a named smart‑marker range, convert column‑wise source data to rows, bind the resulting DataTable to WorkbookDesigner, and generate the final Excel workbook.
// Common Searches: asp.net transpose column data for smart markers Aspose.Cells | process only a specific smart marker range with WorkbookDesigner | how to use Range.Transpose with smart markers in Aspose.Cells | convert column‑wise raw data to rows for smart marker import .NET | named range smart markers example Aspose.Cells
// Tags: range transpose Aspose.Cells | named smart marker range processing | WorkbookDesigner DataTable data source | smart marker import transposed data | excel generation with smart markers .NET

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerTransposeDemo
{
    // The example demonstrates creating a workbook, defining a horizontal smart‑marker layout, naming the marker range, filling raw data column‑wise, transposing that range to row‑wise records, exporting it to a DataTable, binding it to WorkbookDesigner, processing only the named smart‑marker range, and saving the resulting Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Set up smart markers in a horizontal layout (one row will be repeated per data item)
                cells["A2"].PutValue("&=$Data.Name");
                cells["B2"].PutValue("&=$Data.Age");
                cells["C2"].PutValue("&=$Data.City");

                // Name the range that holds the smart markers (required for processing)
                Aspose.Cells.Range smartMarkerRange = cells.CreateRange("A2:C2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // 3. Prepare raw data in a transposed (column‑wise) form.
                cells["E1"].PutValue("John");
                cells["E2"].PutValue(30);
                cells["E3"].PutValue("New York");

                cells["F1"].PutValue("Jane");
                cells["F2"].PutValue(25);
                cells["F3"].PutValue("London");

                // Define the range that contains the raw column‑wise data
                Aspose.Cells.Range rawDataRange = cells.CreateRange("E1:F3");

                // 4. Transpose the raw data so that rows become records (the format expected by smart markers)
                rawDataRange.Transpose();

                // 5. Export the transposed range to a DataTable
                DataTable dataTable = rawDataRange.ExportDataTable();

                // 6. Set up the WorkbookDesigner, assign the DataTable as the data source,
                //    and process only the smart‑marker range.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", dataTable);
                designer.Process(smartMarkerRange, true);

                // 7. Save the resulting workbook
                string outputPath = "SmartMarkerTransposeResult.xlsx";
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
