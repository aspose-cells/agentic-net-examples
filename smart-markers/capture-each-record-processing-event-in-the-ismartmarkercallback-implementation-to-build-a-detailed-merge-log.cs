// Title: Log each smart‑marker record during a WorkbookDesigner merge using ISmartMarkerCallBack in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells how to capture smart marker processing events with ISmartMarkerCallBack | example of logging smart marker merge details to a text file using WorkbookDesigner | record‑level merge log for smart markers in Aspose.Cells .NET | write custom ISmartMarkerCallBack to output sheet row column information | save smart marker processing log as UTF‑8 file in an Aspose.Cells application
// Tags: ISmartMarkerCallBack merge logging C# | smart marker processing callback Aspose.Cells | record level smart marker event capture | write smart marker merge log to UTF-8 file | WorkbookDesigner custom callback example

using System;
using System.Data;
using System.Text;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerMergeLogger
{
    // Implementation of ISmartMarkerCallBack to capture processing events
    // The sample implements ISmartMarkerCallBack to record sheet index, row index, column index, table name, and column name for every smart marker processed by WorkbookDesigner. Log entries are accumulated in a StringBuilder, exposed via GetLog(), and can be written to a UTF‑8 text file. The example shows creating a template workbook, populating a DataTable, assigning the custom MergeLogger as the designer callback, processing the smart markers, saving the merged workbook, and outputting the detailed merge log.
    public class MergeLogger : ISmartMarkerCallBack
    {
        private readonly StringBuilder _logBuilder = new StringBuilder();

        // This method is called by WorkbookDesigner for each smart marker record processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a detailed log entry
            string entry = $"Processed Record - SheetIndex: {sheetIndex}, RowIndex: {rowIndex}, ColumnIndex: {colIndex}, Table: \"{tableName}\", Column: \"{columnName}\"";
            _logBuilder.AppendLine(entry);
        }

        // Retrieve the accumulated log as a string
        public string GetLog()
        {
            return _logBuilder.ToString();
        }

        // Optionally write the log to a file
        public void WriteLogToFile(string filePath)
        {
            File.WriteAllText(filePath, GetLog(), Encoding.UTF8);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Initialize the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner();

            // Create a new workbook (template) in memory
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert smart markers into the worksheet.
            // The syntax "&=$Table1.Column1" tells Aspose.Cells to replace this cell with data from Table1.Column1.
            sheet.Cells["A1"].PutValue("&=$Table1.Column1");
            sheet.Cells["A2"].PutValue("&=$Table1.Column2");
            sheet.Cells["A3"].PutValue("&=$Table1.Column3");

            // Assign the workbook to the designer
            designer.Workbook = workbook;

            // Prepare a DataTable as the data source
            DataTable dataTable = new DataTable("Table1");
            dataTable.Columns.Add("Column1", typeof(string));
            dataTable.Columns.Add("Column2", typeof(string));
            dataTable.Columns.Add("Column3", typeof(string));

            // Add sample rows
            dataTable.Rows.Add("Row1-Value1", "Row1-Value2", "Row1-Value3");
            dataTable.Rows.Add("Row2-Value1", "Row2-Value2", "Row2-Value3");
            dataTable.Rows.Add("Row3-Value1", "Row3-Value2", "Row3-Value3");

            // Set the data source for the designer
            designer.SetDataSource(dataTable);

            // Create and assign the callback logger
            MergeLogger logger = new MergeLogger();
            designer.CallBack = logger;

            // Process the smart markers. 'false' means unrecognized markers will not be preserved.
            designer.Process(false);

            // Save the processed workbook
            string outputWorkbookPath = "MergedOutput.xlsx";
            designer.Workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to \"{outputWorkbookPath}\".");

            // Output the detailed merge log to console
            Console.WriteLine("\n--- Smart Marker Merge Log ---");
            Console.WriteLine(logger.GetLog());

            // Optionally write the log to a text file
            string logFilePath = "merge_log.txt";
            logger.WriteLogToFile(logFilePath);
            Console.WriteLine($"Merge log written to \"{logFilePath}\".");
        }
    }
}
