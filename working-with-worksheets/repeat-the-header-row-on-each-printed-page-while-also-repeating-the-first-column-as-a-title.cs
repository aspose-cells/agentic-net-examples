// Title: C# – Repeat Header Row & First Column on Every Printed Page with Aspose.Cells
// Description: Creates a new Workbook, adds a header row (A1:C1) and sample data (rows 2‑50), then uses Worksheet.PageSetup to set PrintTitleRows = "$1:$1" and PrintTitleColumns = "$A:$A" so the first row and column repeat on each printed page. Saves as RepeatHeaderAndFirstColumn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | repeat header row | repeat first column | PrintTitleRows | PrintTitleColumns | PageSetup | print titles | Excel pagination | multi‑page print | Aspose.Cells example | global | US
// Common Searches: Aspose.Cells repeat header row on each printed page | How to set first column as title in Aspose.Cells .NET | PrintTitleRows and PrintTitleColumns Aspose.Cells example | C# code to repeat rows and columns when printing Excel | Aspose.Cells page setup for multi‑page reports
// Developer Intent: Set up a worksheet so the top row and leftmost column are printed on every page of a multi‑page Excel document.
// Use Cases: Printing large tables where column A holds identifiers that must appear on every page | Generating multi‑page invoices or reports with persistent column headers and row titles | Creating printable schedules or timetables that need both header row and title column repeated
// AI Prompts: Provide C# code that configures Aspose.Cells PageSetup to repeat the first row and first column on each printed page. | Explain how to test that PrintTitleRows and PrintTitleColumns are applied correctly in an Aspose.Cells workbook. | Show an Aspose.Cells example that repeats multiple header rows and a title column when printing.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, adds a header row (A1:C1) and sample data (rows 2‑50), then uses Worksheet.PageSetup to set PrintTitleRows = "$1:$1" and PrintTitleColumns = "$A:$A" so the first row and column repeat on each printed page. Saves as RepeatHeaderAndFirstColumn.xlsx.
    public class RepeatHeaderAndFirstColumnDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data: header row
            worksheet.Cells["A1"].PutValue("Title");
            worksheet.Cells["B1"].PutValue("Header1");
            worksheet.Cells["C1"].PutValue("Header2");

            // Sample data: multiple rows to cause pagination
            for (int i = 2; i <= 50; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue($"Data {i - 1} - 1");
                worksheet.Cells[$"C{i}"].PutValue($"Data {i - 1} - 2");
            }

            // Configure page setup to repeat the first row and first column on each printed page
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.PrintTitleRows = "$1:$1";   // repeat row 1 (header)
            pageSetup.PrintTitleColumns = "$A:$A"; // repeat column A (title)

            // Save the workbook
            workbook.Save("RepeatHeaderAndFirstColumn.xlsx");
        }
    }
}
