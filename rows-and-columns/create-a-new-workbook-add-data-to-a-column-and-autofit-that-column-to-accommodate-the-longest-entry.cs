// Title: C# Aspose.Cells example: create a workbook, add variable‑length text to column A, and auto‑fit the column width
// AI Prompts: Write C# code using Aspose.Cells that creates a new workbook, inserts strings of different lengths into column A, calls AutoFitColumn for that column, and saves the file as an .xlsx. | Show how to populate a worksheet column with mixed‑length values and then automatically adjust the column width using worksheet.AutoFitColumn in Aspose.Cells for .NET. | Demonstrate the steps to generate a workbook, write data to a specific column, apply auto‑fit to that column only, and persist the workbook to disk with Aspose.Cells C# API.
// Common Searches: aspnet c# how to automatically adjust column width after inserting data with Aspose.Cells | example of worksheet.AutoFitColumn for a single column in Aspose.Cells .NET | save workbook as xlsx after column auto‑fit using Aspose.Cells C# | populate column A with strings of varying length and auto‑size column in Aspose.Cells | Aspose.Cells C# auto‑fit column only after adding rows
// Tags: auto-fit column Aspose.Cells C# | write variable length strings to worksheet column Aspose.Cells | worksheet.AutoFitColumn single column example | save workbook as xlsx Aspose.Cells .NET | create new workbook Aspose.Cells C# | adjust column width based on content Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsColumnAutoFitDemo
{
    // The program creates a new workbook, writes several strings of varying lengths to column A of the first worksheet, auto‑fits column A to the longest entry, and saves the workbook as ColumnAutoFitDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to column A (index 0) with varying lengths
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["A2"].PutValue("Medium length text");
            worksheet.Cells["A3"].PutValue("This is a much longer piece of text that should cause the column to expand");
            worksheet.Cells["A4"].PutValue("Tiny");
            worksheet.Cells["A5"].PutValue("Another long text entry to demonstrate auto‑fit functionality");

            // Auto‑fit only column A (column index 0)
            worksheet.AutoFitColumn(0);

            // Save the workbook to a file
            workbook.Save("ColumnAutoFitDemo.xlsx");
        }
    }
}
