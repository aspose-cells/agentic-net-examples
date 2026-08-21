// Title: AutoFitColumns: Automatically size worksheet columns in Aspose.Cells for .NET (C#)
// Description: C# sample that creates a workbook, writes short, medium, and long text strings to cells A1‑C1, invokes Worksheet.AutoFitColumns() to resize every column to its content, and saves the result as AutoFitColumnsResult.xlsx.
// Keywords: Aspose.Cells AutoFitColumns | C# column autosize | adjust Excel column width programmatically | auto size columns .NET | worksheet.AutoFitColumns method | Excel column width based on content
// Common Searches: Aspose.Cells AutoFitColumns C# | how to auto size Excel columns using Aspose.Cells | C# auto fit columns worksheet | auto adjust column width Aspose.Cells .NET | auto fit all columns Aspose.Cells example
// Developer Intent: Resize every column in a worksheet so it fits the longest cell value.
// Use Cases: Generate a dynamic report where column widths must adapt to varying text lengths before exporting to Excel. | Populate a template with data rows and automatically polish the layout by fitting columns to content. | Create an Excel file from user‑generated data and ensure a professional appearance without manual width adjustments.
// AI Prompts: Show how to auto‑fit columns for a specific range (e.g., A1:C10) instead of the entire worksheet using Aspose.Cells. | Provide C# code that adds rows in a loop and calls AutoFitColumns only after all data is written. | Explain how to combine Worksheet.AutoFitColumns with a maximum column width limit in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AutoFitColumnsDemo
{
    // C# sample that creates a workbook, writes short, medium, and long text strings to cells A1‑C1, invokes Worksheet.AutoFitColumns() to resize every column to its content, and saves the result as AutoFitColumnsResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to demonstrate column width adjustment
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["B1"].PutValue("Medium length text");
            worksheet.Cells["C1"].PutValue("This is a very long text that should cause the column to expand automatically");

            // Auto-fit all columns in the worksheet to match the content
            worksheet.AutoFitColumns();

            // Save the workbook (save rule)
            workbook.Save("AutoFitColumnsResult.xlsx");
        }
    }
}
