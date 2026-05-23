using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Logger for smart‑marker processing
    public class SmartMarkerAuditLogger : ISmartMarkerCallBack
    {
        private readonly List<string> _log = new List<string>();

        // Called for each smart‑marker replacement
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}";
            _log.Add(entry);
            Console.WriteLine(entry);
        }

        public IEnumerable<string> GetLog() => _log;

        // Save log to a text file
        public void SaveLog(string filePath)
        {
            File.WriteAllLines(filePath, _log);
        }
    }

    public class SmartMarkerProcessor
    {
        public static void Run()
        {
            try
            {
                const string templatePath = "template.xlsx";

                // Ensure template exists; create a minimal one if missing
                if (!File.Exists(templatePath))
                {
                    var wb = new Workbook();
                    wb.Worksheets[0].Name = "Sheet1";
                    // Example smart marker (optional)
                    wb.Worksheets[0].Cells["A1"].PutValue("&=Products.ProductName");
                    wb.Save(templatePath);
                }

                // Load the template workbook
                Workbook template = new Workbook(templatePath);

                // Initialize designer with logger callback
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = template,
                    CallBack = new SmartMarkerAuditLogger()
                };

                // Sample data source
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("ProductName", typeof(string));
                dt.Columns.Add("Price", typeof(double));
                dt.Rows.Add("Apple", 1.2);
                dt.Rows.Add("Banana", 0.8);

                designer.SetDataSource(dt);

                // Process smart markers (do not preserve unrecognized markers)
                designer.Process(false);

                // Save the processed workbook
                const string outputPath = "output.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");

                // Persist audit log
                if (designer.CallBack is SmartMarkerAuditLogger logger)
                {
                    const string logPath = "SmartMarkerAuditLog.txt";
                    logger.SaveLog(logPath);
                    Console.WriteLine($"Audit log saved to {logPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point
    public static class Program
    {
        public static void Main()
        {
            SmartMarkerProcessor.Run();
        }
    }
}