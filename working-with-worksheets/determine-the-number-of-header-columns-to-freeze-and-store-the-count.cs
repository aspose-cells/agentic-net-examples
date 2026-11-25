// Title: Detect and Freeze Consecutive Header Columns in an Excel Worksheet and Save the Count as a Custom Property using Aspose.Cells for .NET
// Common Searches: C# count non‑empty header cells in first row and freeze those columns with Aspose.Cells | how to store number of frozen columns in a custom document property using Aspose.Cells .NET | dynamic FreezePanes based on header row length Aspose.Cells C# example | Aspose.Cells detect header columns until empty cell and apply freeze panes | save custom property HeaderColumnsFrozen in Excel workbook with Aspose.Cells
// Tags: detect header column count Aspose.Cells | freeze panes based on header length C# | custom document property for frozen columns Aspose.Cells | dynamic column freeze Excel .NET | header row analysis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace FreezeHeaderColumnsDemo
{
    // The sample loads an Excel workbook, iterates across the first row to count consecutive non‑empty header cells, freezes that many columns with FreezePanes, records the count in a custom document property named "HeaderColumnsFrozen", and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook. Adjust as needed.
            string inputPath = "HeaderDemo.xlsx";

            // Ensure the Aspose.Cells license is set if you have one.
            // License license = new License();
            // license.SetLicense("Aspose.Total.NET.lic");

            // If the file does not exist, create a simple workbook with sample headers.
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbook(inputPath);
                Console.WriteLine($"Sample workbook created at '{Path.GetFullPath(inputPath)}'.");
            }

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to load workbook '{inputPath}': {ex.Message}");
                return;
            }

            // Access the first worksheet (index 0).
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the number of header columns.
            // Header columns are defined as consecutive non‑empty cells starting from column A in row 0.
            int headerColumnCount = 0;
            Cells cells = sheet.Cells;
            while (true)
            {
                Cell cell = cells[0, headerColumnCount]; // Row 0, current column
                if (cell == null || string.IsNullOrWhiteSpace(cell.StringValue))
                {
                    // Stop at the first empty cell.
                    break;
                }
                headerColumnCount++;
            }

            Console.WriteLine($"Detected {headerColumnCount} header column(s) to freeze.");

            // Freeze the determined number of columns.
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // We freeze rows = 0 (no frozen rows) and columns = headerColumnCount.
            // The freeze position is the cell just after the frozen area, i.e., column index = headerColumnCount.
            // Row index is 0 because we are not freezing any rows.
            if (headerColumnCount > 0)
            {
                // The cell at (0, headerColumnCount) is the first unfrozen cell.
                sheet.FreezePanes(0, headerColumnCount, 0, headerColumnCount);
                Console.WriteLine($"Frozen the first {headerColumnCount} column(s).");
            }
            else
            {
                Console.WriteLine("No header columns found; nothing to freeze.");
            }

            // Store the count in a custom document property for persistence.
            const string propertyName = "HeaderColumnsFrozen";
            try
            {
                // Remove existing property if it exists.
                if (workbook.CustomDocumentProperties[propertyName] != null)
                {
                    workbook.CustomDocumentProperties.Remove(propertyName);
                }

                // Add the new property.
                workbook.CustomDocumentProperties.Add(propertyName, headerColumnCount);
                Console.WriteLine($"Stored header column count ({headerColumnCount}) in custom property '{propertyName}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to store custom property: {ex.Message}");
            }

            // Save the workbook with the freeze applied.
            string outputPath = "HeaderDemo_Frozen.xlsx";
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with frozen panes to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a simple workbook with a header row containing sample data.
        /// </summary>
        /// <param name="path">File path where the workbook will be saved.</param>
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Sample header values (you can modify as needed).
            string[] headers = { "ID", "Name", "Date", "Amount", "Status" };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[0, i].PutValue(headers[i]);
            }

            // Add a few data rows for completeness.
            ws.Cells[1, 0].PutValue(1);
            ws.Cells[1, 1].PutValue("Alice");
            ws.Cells[1, 2].PutValue(DateTime.Today);
            ws.Cells[1, 3].PutValue(123.45);
            ws.Cells[1, 4].PutValue("Open");

            ws.Cells[2, 0].PutValue(2);
            ws.Cells[2, 1].PutValue("Bob");
            ws.Cells[2, 2].PutValue(DateTime.Today.AddDays(-1));
            ws.Cells[2, 3].PutValue(678.90);
            ws.Cells[2, 4].PutValue("Closed");

            wb.Save(path, SaveFormat.Xlsx);
        }
    }
}
