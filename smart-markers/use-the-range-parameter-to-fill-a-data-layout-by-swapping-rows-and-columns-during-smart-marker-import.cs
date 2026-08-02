// Title: Transpose Smart Marker Range to Swap Rows and Columns in Aspose.Cells (C#)
// Description: Demonstrates how to create a named range containing vertical smart markers, use Range.Transpose to flip rows and columns, process only the transposed range with WorkbookDesigner, and save the workbook. Ideal for converting a column‑wise smart‑marker layout into a row‑wise table in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# smart markers | Range.Transpose | named range smart markers | swap rows and columns | smart marker layout | WorkbookDesigner.Process | transpose smart markers | Aspose.Cells .NET
// Common Searches: How to use Range.Transpose with smart markers in Aspose.Cells | Swap rows and columns of smart markers in C# Aspose.Cells | Process a specific smart marker range after transposition | Create and name a smart marker range before transposing | Aspose.Cells transpose vertical smart markers to horizontal
// Developer Intent: Change the orientation of a smart‑marker range by transposing rows and columns before processing it with WorkbookDesigner.
// Use Cases: Convert a vertical list of smart markers into a horizontal table for compact reporting | Generate alternative layouts (row‑wise vs column‑wise) from the same data source | Apply smart‑marker processing only to a selected range while preserving other worksheet content | Reuse the same data source for multiple transposed layouts in a single workbook
// AI Prompts: Generate C# code that defines vertical smart markers, creates a named range, calls Range.Transpose, and processes only that range with WorkbookDesigner. | Explain how Range.Transpose affects smart‑marker placeholders and how to revert the layout after processing. | Provide a step‑by‑step tutorial for binding a DataTable to smart markers, transposing the marker range, and saving the workbook with the new layout.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerTransposeDemo
{
    // Demonstrates how to create a named range containing vertical smart markers, use Range.Transpose to flip rows and columns, process only the transposed range with WorkbookDesigner, and save the workbook. Ideal for converting a column‑wise smart‑marker layout into a row‑wise table in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // ------------------------------------------------------------
                // 1. Define smart markers in a vertical layout (two rows, one column)
                // ------------------------------------------------------------
                cells["A2"].PutValue("&=$Data.Month");
                cells["A3"].PutValue("&=$Data.Sales");

                // ------------------------------------------------------------
                // 2. Prepare the data source (a DataTable with two columns)
                // ------------------------------------------------------------
                DataTable dt = new DataTable();
                dt.Columns.Add("Month", typeof(string));
                dt.Columns.Add("Sales", typeof(double));
                dt.Rows.Add("Jan", 1200.5);
                dt.Rows.Add("Feb", 950.75);
                dt.Rows.Add("Mar", 1340.0);

                // ------------------------------------------------------------
                // 3. Set up WorkbookDesigner and bind the data source
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", dt);

                // ------------------------------------------------------------
                // 4. Create a Range that contains the smart markers
                // ------------------------------------------------------------
                Aspose.Cells.Range smartRange = cells.CreateRange("A2:A3");
                smartRange.Name = "_CellsSmartMarkers";

                // ------------------------------------------------------------
                // 5. Transpose the range to swap rows and columns
                // ------------------------------------------------------------
                smartRange.Transpose();

                // ------------------------------------------------------------
                // 6. Process only the transposed range
                // ------------------------------------------------------------
                designer.Process(smartRange, true);

                // ------------------------------------------------------------
                // 7. Save the result
                // ------------------------------------------------------------
                string outputPath = "SmartMarkerTransposeResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
