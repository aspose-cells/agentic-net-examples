// Title: Log Smart Marker Merges Using ISmartMarkerCallBack in Aspose.Cells for .NET
// Description: Shows how to implement ISmartMarkerCallBack to capture every smart‑marker processing event during WorkbookDesigner merging. The callback records sheet, row, column, table and column names in a List<string>, outputs the log to the console, saves it to a text file, and writes the merged workbook. Includes a helper that creates a minimal template when none is present.
// Keywords: Aspose.Cells | ISmartMarkerCallBack | smart marker log | C# example | WorkbookDesigner | merge audit | data merge debugging | Excel template | smart markers .NET | record merge events
// Common Searches: How to log each smart marker replacement with Aspose.Cells | ISmartMarkerCallBack example for merge auditing in C# | Capture smart marker processing details during WorkbookDesigner merge | Aspose.Cells merge log for smart markers | Debug smart marker data merge in .NET
// Developer Intent: Create a detailed audit trail of every smart‑marker replacement while merging data with Aspose.Cells.
// Use Cases: Debugging incorrect data placement by reviewing sheet, row, and column logs | Compliance auditing of Excel data merges with a chronological record | Performance analysis of large smart‑marker merges | Generating custom reports of processed tables and columns | Automated unit tests that verify expected smart‑marker processing
// AI Prompts: Write a C# extension for MergeLogSmartMarkerCallback that adds a UTC timestamp to each entry and exports the log as JSON. | Show how to filter the merge log to include only entries for a specific table name when using ISmartMarkerCallBack. | Provide code to read MergeLog.txt and display the entries in a Windows Forms ListView with sortable columns.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerMergeLogDemo
{
    // Callback implementation that records each smart marker processing event
    // Shows how to implement ISmartMarkerCallBack to capture every smart‑marker processing event during WorkbookDesigner merging. The callback records sheet, row, column, table and column names in a List<string>, outputs the log to the console, saves it to a text file, and writes the merged workbook. Includes a helper that creates a minimal template when none is present.
    public class MergeLogSmartMarkerCallback : ISmartMarkerCallBack
    {
        // Stores detailed log entries
        public List<string> Log { get; } = new List<string>();

        // Called by Aspose.Cells for each smart marker occurrence
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a descriptive log entry
            string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:\"{tableName}\", Column:\"{columnName}\"";
            Log.Add(entry);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template workbook containing smart markers
            string templatePath = "Template.xlsx";

            // Ensure the template file exists (for demo purposes we create a simple one if missing)
            if (!File.Exists(templatePath))
            {
                CreateSampleTemplate(templatePath);
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Create and assign the custom callback
            MergeLogSmartMarkerCallback callback = new MergeLogSmartMarkerCallback();
            designer.CallBack = callback;

            // Prepare a simple data source matching the smart markers in the template
            DataTable data = new DataTable("Products");
            data.Columns.Add("ProductName", typeof(string));
            data.Columns.Add("Price", typeof(double));
            data.Rows.Add("Apple", 1.2);
            data.Rows.Add("Banana", 0.8);
            data.Rows.Add("Cherry", 2.5);

            // Set the data source for the designer
            designer.SetDataSource(data);

            // Process the smart markers (true = preserve unrecognized markers)
            designer.Process(true);

            // Save the processed workbook
            string outputPath = "ProcessedOutput.xlsx";
            designer.Workbook.Save(outputPath);

            // Output the merge log to console
            Console.WriteLine("Smart Marker Processing Log:");
            foreach (string logEntry in callback.Log)
            {
                Console.WriteLine(logEntry);
            }

            // Optionally, write the log to a text file
            string logFilePath = "MergeLog.txt";
            File.WriteAllLines(logFilePath, callback.Log);
            Console.WriteLine($"Merge log saved to: {logFilePath}");
        }

        // Helper method to create a minimal template with smart markers if none exists
        private static void CreateSampleTemplate(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            // Insert smart markers that reference the "Products" table
            ws.Cells["A1"].PutValue("&=Products.ProductName");
            ws.Cells["B1"].PutValue("&=Products.Price");
            // Name the range for smart markers (required when LineByLine = false, but not needed here)
            ws.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";
            wb.Save(path);
        }
    }
}
