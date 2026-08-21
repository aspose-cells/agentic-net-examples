// Title: Aspose.Cells C# – Set Print Title Columns A:B for Repeating Headers
// Description: Shows how to assign the PageSetup.PrintTitleColumns property to "$A:$B" so that columns A and B appear on every printed page, adds sample rows, and saves the workbook as PrintTitleColumns_AB.xlsx.
// Keywords: Aspose.Cells | PrintTitleColumns | repeat columns | C# .NET | worksheet print titles | Excel export | PageSetup | column headers | Aspose.Cells example | set print titles
// Common Searches: Aspose.Cells set PrintTitleColumns C# | repeat columns A B on printed pages Aspose | how to make column headers repeat in Excel using Aspose.Cells | PageSetup PrintTitleColumns property example | C# code to set repeating columns in workbook
// Developer Intent: Configure columns A and B as repeating print titles for a worksheet using Aspose.Cells for .NET.
// Use Cases: Create printable reports where identifier columns stay visible on each page. | Export data to Excel with persistent header columns for multi‑page printing. | Generate invoices or catalogs that require the first two columns to repeat across pages.
// AI Prompts: Provide a C# snippet that sets columns A through C as print title columns with Aspose.Cells and saves the file. | Explain how to configure both print title rows and columns together using the PageSetup object. | Show how to read the current PrintTitleColumns value from a worksheet, modify it, and persist the change.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to assign the PageSetup.PrintTitleColumns property to "$A:$B" so that columns A and B appear on every printed page, adds sample rows, and saves the workbook as PrintTitleColumns_AB.xlsx.
    public class SetPrintTitleColumnsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the columns to repeat on the left side of each printed page (columns A and B)
                worksheet.PageSetup.PrintTitleColumns = "$A:$B";

                // Optional: add some data to demonstrate the repeated columns
                for (int row = 1; row <= 50; row++)
                {
                    worksheet.Cells[$"A{row}"].PutValue($"Header A - Row {row}");
                    worksheet.Cells[$"B{row}"].PutValue($"Header B - Row {row}");
                    worksheet.Cells[$"C{row}"].PutValue($"Data {row}");
                }

                // Save the workbook
                workbook.Save("PrintTitleColumns_AB.xlsx");
                Console.WriteLine("Workbook saved successfully as PrintTitleColumns_AB.xlsx");
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
            SetPrintTitleColumnsDemo.Run();
        }
    }
}
