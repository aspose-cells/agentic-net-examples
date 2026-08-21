// Title: C# – Repeat Columns A and B on Every Printed Page with Aspose.Cells
// Description: Demonstrates how to set the PrintTitleColumns property to "$A:$B" so columns A and B appear on each printed page, adds sample data, and saves the workbook as PrintTitleColumnsAB.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PrintTitleColumns | repeat columns on printed page C# | Aspose.Cells set print titles | C# Excel repeat columns | Aspose.Cells workbook printing
// Common Searches: Aspose.Cells repeat columns A B each page | C# set PrintTitleColumns property | how to print title columns with Aspose.Cells | Aspose.Cells repeat column headers on print
// Developer Intent: Configure columns A and B to repeat as title columns on every printed page of an Excel workbook.
// Use Cases: Multi‑page reports that need header columns visible on each page. | Invoices or catalogs where product description columns must appear on every printed sheet. | Large data exports where column identifiers should stay on each printed page for readability.
// AI Prompts: Show how to set non‑adjacent columns as print titles with Aspose.Cells in C#. | Provide example code to configure both print title rows and columns together. | Explain how to programmatically verify the PrintTitleColumns setting after saving the file.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set the PrintTitleColumns property to "$A:$B" so columns A and B appear on each printed page, adds sample data, and saves the workbook as PrintTitleColumnsAB.xlsx using Aspose.Cells for .NET.
    public class PrintTitleColumnsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set columns A and B to repeat on each printed page
                worksheet.PageSetup.PrintTitleColumns = "$A:$B";

                // Add sample data to illustrate the effect
                for (int i = 0; i < 100; i++)
                {
                    worksheet.Cells[i, 0].PutValue($"Row {i + 1} - Column A");
                    worksheet.Cells[i, 1].PutValue($"Row {i + 1} - Column B");
                    worksheet.Cells[i, 2].PutValue($"Row {i + 1} - Column C");
                }

                // Save the workbook
                workbook.Save("PrintTitleColumnsAB.xlsx");
                Console.WriteLine("Workbook saved successfully as PrintTitleColumnsAB.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PrintTitleColumnsDemo.Run();
        }
    }
}
