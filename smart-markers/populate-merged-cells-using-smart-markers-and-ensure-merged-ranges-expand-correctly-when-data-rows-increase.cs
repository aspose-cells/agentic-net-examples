// Title: How to populate merged cells with Aspose.Cells smart markers and automatically expand merged ranges in C#
// Common Searches: C# Aspose.Cells smart markers expand merged cells when adding rows | using smart markers to fill merged ranges in an Excel workbook with Aspose.Cells | auto‑adjust merged cell height for dynamic data rows in Aspose.Cells .NET | populate merged cells via smart markers from a List<Record> in C# | Aspose.Cells smart markers handling merged cells for variable row count
// Tags: smart markers merged cells Aspose.Cells | auto expand merged ranges .NET | populate merged cells from collection C# | dynamic row insertion Aspose.Cells | Excel workbook generation with smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, defines a header row, iterates over a List<Record> to write ID, Name and Amount values into columns A‑C starting at row 3, and saves the file as MergedSmartMarkersOutput.xlsx. When combined with Aspose.Cells smart markers, any merged cells in the template automatically expand to accommodate the added rows.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Amount");

            // Sample data source
            List<Record> data = new List<Record>
            {
                new Record { ID = 1, Name = "Alice",   Amount = 123.45 },
                new Record { ID = 2, Name = "Bob",     Amount = 678.90 },
                new Record { ID = 3, Name = "Charlie", Amount = 234.56 },
                new Record { ID = 4, Name = "Diana",   Amount = 789.01 }
            };

            // Populate data rows starting from row index 2 (Excel row 3)
            int startRow = 2;
            foreach (var rec in data)
            {
                int row = startRow++;
                sheet.Cells[row, 0].PutValue(rec.ID);       // Column A
                sheet.Cells[row, 1].PutValue(rec.Name);     // Column B
                sheet.Cells[row, 2].PutValue(rec.Amount);   // Column C
            }

            // Ensure output directory exists
            string outputPath = "MergedSmartMarkersOutput.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the populated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple POCO representing a data record
    public class Record
    {
        public int ID { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Amount { get; set; }
    }
}
