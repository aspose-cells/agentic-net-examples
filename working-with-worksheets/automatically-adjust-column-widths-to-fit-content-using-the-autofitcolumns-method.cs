// Title: AutoFitColumns – Auto‑size worksheet columns to content with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, writes three strings of different lengths into cells A1, B1, and C1, calls worksheet.AutoFitColumns() to resize every column based on its longest cell, and saves the file as AutoFitColumnsDemo.xlsx.
// Keywords: Aspose.Cells AutoFitColumns | C# auto size columns | adjust column width programmatically | fit column width to content .NET | worksheet.AutoFitColumns example | auto‑fit columns Aspose.Cells
// Common Searches: Aspose.Cells AutoFitColumns C# | how to auto size columns in Excel using Aspose.Cells | auto fit column width .NET | C# code to fit columns to content Aspose | auto adjust column width Aspose.Cells worksheet
// Developer Intent: Programmatically resize worksheet columns so each column width matches the length of its longest cell value.
// Use Cases: Generate a data report where column widths automatically adapt to varying text lengths before saving. | Export a DataTable to Excel and call AutoFitColumns to improve readability without manual adjustments. | Populate a template‑based spreadsheet with dynamic content and ensure columns are sized to prevent truncation. | Create a dashboard workbook where new rows are added daily and columns stay optimally sized.
// AI Prompts: Provide C# code that uses Aspose.Cells to auto‑fit columns for a specific range after populating data. | Explain how to set minimum and maximum column widths while using AutoFitColumns in Aspose.Cells for .NET. | Show an example of auto‑fitting columns after importing a DataTable into a worksheet with Aspose.Cells. | Demonstrate combining AutoFitColumns with column hiding and pane freezing in a C# workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, writes three strings of different lengths into cells A1, B1, and C1, calls worksheet.AutoFitColumns() to resize every column based on its longest cell, and saves the file as AutoFitColumnsDemo.xlsx.
    public class AutoFitColumnsExample
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with sample data of varying lengths
            worksheet.Cells["A1"].PutValue("This is a test string");
            worksheet.Cells["B1"].PutValue("Another longer test string for demonstration");
            worksheet.Cells["C1"].PutValue("Short");

            // Auto-fit all columns in the worksheet to match the content width
            worksheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("AutoFitColumnsDemo.xlsx");
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                AutoFitColumnsExample.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
