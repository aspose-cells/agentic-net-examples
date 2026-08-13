// Title: Auto‑Fit All Columns in an Excel Worksheet with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, populates a simple table, calls Worksheet.AutoFitColumns() to size every column to its longest cell value, and saves the result as AutoFitAllColumnsDemo.xlsx.
// Keywords: Aspose.Cells AutoFitColumns | C# Excel column autosize | adjust column width Aspose.Cells | Worksheet.AutoFitColumns example | .NET Excel column auto‑fit
// Common Searches: Aspose.Cells auto‑fit all columns C# | Worksheet.AutoFitColumns usage example | how to size Excel columns to content with Aspose.Cells | C# code to auto‑size columns in a workbook | Aspose.Cells column width adjustment
// Developer Intent: Resize every column in a worksheet so its width matches the longest cell value automatically.
// Use Cases: Generate Excel reports where column widths adapt to varying text lengths. | Export data tables without manually setting column sizes, ensuring a clean layout. | Prepare spreadsheets for printing or sharing with optimal column widths for readability.
// AI Prompts: Show how to auto‑fit columns for a specific range instead of the entire sheet using Aspose.Cells. | Provide C# code to auto‑fit rows after adjusting column widths with Aspose.Cells. | Explain how to limit the maximum column width when using Worksheet.AutoFitColumns.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, populates a simple table, calls Worksheet.AutoFitColumns() to size every column to its longest cell value, and saves the result as AutoFitAllColumnsDemo.xlsx.
    public class AutoFitAllColumnsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data that represents a table
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["C1"].PutValue("Description");

                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("Alice");
                worksheet.Cells["C2"].PutValue("Short description");

                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Bob");
                worksheet.Cells["C3"].PutValue("This is a much longer description that should cause the column to expand.");

                worksheet.Cells["A4"].PutValue(3);
                worksheet.Cells["B4"].PutValue("Charlie");
                worksheet.Cells["C4"].PutValue("Medium length text");

                // Auto‑fit all columns so each column width matches its longest cell content
                worksheet.AutoFitColumns();

                // Save the workbook
                workbook.Save("AutoFitAllColumnsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutoFitAllColumnsDemo.Run();
        }
    }
}
