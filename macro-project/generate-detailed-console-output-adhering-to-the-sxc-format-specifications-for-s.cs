using System;
using Aspose.Cells;

namespace AsposeCellsSxcLoggingDemo
{
    enum LogLevel
    {
        Info,
        Error
    }

    class Program
    {
        static void Log(LogLevel level, string message)
        {
            Console.WriteLine($"{{\"level\":\"{level}\",\"timestamp\":\"{DateTime.UtcNow:O}\",\"message\":\"{message}\"}}");
        }

        static void Main()
        {
            try
            {
                Log(LogLevel.Info, "Creating a new workbook.");
                Workbook workbook = new Workbook();

                Log(LogLevel.Info, "Accessing the first worksheet.");
                Worksheet sheet = workbook.Worksheets[0];

                Log(LogLevel.Info, "Populating sample data.");
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(92);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");
                sheet.Cells["C4"].PutValue(78);

                Log(LogLevel.Info, "Saving workbook in SXC format.");
                string outputPath = "StructuredDataDemo.sxc";
                workbook.Save(outputPath, SaveFormat.Sxc);

                Log(LogLevel.Info, $"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"An exception occurred: {ex.Message}");
            }
        }
    }
}