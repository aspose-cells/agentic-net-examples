// Title: Aspose.Cells .NET: Insert Smart Marker with Array Index Placeholder and Log Processing via ISmartMarkerCallBack
// Description: This example creates a workbook, writes a smart‑marker string with an array index placeholder (e.g., &Data[${row}]) into cell A1, binds a List<string> as the "Data" source, attaches an ISmartMarkerCallBack implementation to log sheet, row, column, table and column names, processes the markers with WorkbookDesigner, and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | Smart Markers | Array index placeholder | ISmartMarkerCallBack | WorkbookDesigner | callback logging | Excel generation | GitHub example
// Common Searches: Aspose.Cells smart marker array index placeholder | How to use ISmartMarkerCallBack in Aspose.Cells | Place &Data[${row}] marker in Excel with Aspose | Log smart marker processing in C# | Aspose.Cells example for smart markers on GitHub
// Developer Intent: Add a smart marker containing an array index placeholder to a worksheet cell, bind a matching data source, and capture each marker's processing details through a callback.
// Use Cases: Generate rows from a List<string> by using &Data[${row}] in a cell. | Track smart‑marker resolution (sheet, row, column, table, column) for debugging or auditing. | Create reusable Excel templates where placeholders are replaced at runtime. | Integrate smart‑marker processing into automated report pipelines.
// AI Prompts: Show how to replace the ${row} placeholder with ${col} in the smart marker. | Provide C# code that writes the callback output to a log file instead of the console. | Explain how to bind a DataTable to the marker "&Data[${row}]" and retrieve the filled values after designer.Process().

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Callback implementation to capture smart marker processing details
    // This example creates a workbook, writes a smart‑marker string with an array index placeholder (e.g., &Data[${row}]) into cell A1, binds a List<string> as the "Data" source, attaches an ISmartMarkerCallBack implementation to log sheet, row, column, table and column names, processes the markers with WorkbookDesigner, and saves the result as an Excel file.
    public class IndexPlaceholderCallback : ISmartMarkerCallBack
    {
        // This method is invoked for each smart marker occurrence
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Output the indices and marker information to the console
            Console.WriteLine($"SmartMarker processed - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle rule: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 2. Place a marker string with array index placeholders in a cell
            //    Example marker: &Data[${row}] (the placeholder ${row} will be replaced by the actual row index)
            // -------------------------------------------------
            // Cell A1 will contain the smart marker
            sheet.Cells["A1"].PutValue("&Data[${row}]");

            // -------------------------------------------------
            // 3. Prepare a simple data source (list of strings)
            // -------------------------------------------------
            List<string> data = new List<string> { "Alpha", "Beta", "Gamma" };
            // The data source name must match the table name used in the marker ("Data")
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign the callback that will be invoked during processing
                CallBack = new IndexPlaceholderCallback()
            };
            designer.SetDataSource("Data", data);

            // -------------------------------------------------
            // 4. Process the smart markers (the marker placed in step 2 will be resolved)
            // -------------------------------------------------
            designer.Process();

            // -------------------------------------------------
            // 5. Save the resulting workbook (lifecycle rule: save)
            // -------------------------------------------------
            workbook.Save("SmartMarkerWithIndexPlaceholders.xlsx");

            Console.WriteLine("Workbook saved as SmartMarkerWithIndexPlaceholders.xlsx");
        }
    }
}
