// Title: Repeat Header Row and First Column on Every Printed Page – Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, fills a range (A1:D30) with a title row, column headers and data, then uses PageSetup to set PrintTitleRows = "$1:$1" and PrintTitleColumns = "$A:$A" so the header row and the first column appear on each printed page. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells repeat header row | Aspose.Cells repeat first column | PrintTitleRows C# | PrintTitleColumns .NET | page setup repeat titles | Aspose.Cells printing options | C# workbook print area | Aspose.Cells pagination
// Common Searches: Aspose.Cells how to repeat header row on each printed page | repeat first column in printed Excel using Aspose.Cells C# | set PrintTitleRows and PrintTitleColumns in .NET | configure page setup for repeating titles Aspose.Cells | define print area and repeat titles Aspose.Cells
// Developer Intent: Configure a worksheet so that both the top header row and the leftmost column are printed on every page of a multi‑page document.
// Use Cases: Generating multi‑page reports where column headings must stay visible on each sheet. | Creating printable inventory lists that keep the item name column as a persistent title. | Producing financial statements with a fixed title column and header row for consistent pagination.
// AI Prompts: Show C# code that sets PrintTitleRows and PrintTitleColumns in Aspose.Cells and defines a custom print area. | Provide an example that saves the workbook as PDF after configuring the page setup to repeat both header row and first column. | Explain how to adjust PageSetup properties to repeat titles without altering existing worksheet data.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, fills a range (A1:D30) with a title row, column headers and data, then uses PageSetup to set PrintTitleRows = "$1:$1" and PrintTitleColumns = "$A:$A" so the header row and the first column appear on each printed page. The workbook is saved as an XLSX file.
    public class RepeatHeaderAndFirstColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Title");
                worksheet.Cells["B1"].PutValue("Header 1");
                worksheet.Cells["C1"].PutValue("Header 2");
                worksheet.Cells["D1"].PutValue("Header 3");

                // Data rows (first column will act as a title column)
                for (int row = 2; row <= 30; row++)
                {
                    worksheet.Cells[row - 1, 0].PutValue("Row Title " + (row - 1));
                    worksheet.Cells[row - 1, 1].PutValue("Data " + (row - 1) + "-1");
                    worksheet.Cells[row - 1, 2].PutValue("Data " + (row - 1) + "-2");
                    worksheet.Cells[row - 1, 3].PutValue("Data " + (row - 1) + "-3");
                }

                // Configure page setup to repeat header row and first column
                PageSetup pageSetup = worksheet.PageSetup;
                pageSetup.PrintTitleRows = "$1:$1";
                pageSetup.PrintTitleColumns = "$A:$A";
                pageSetup.PrintArea = "A1:D30";

                // Define output file path
                string outputPath = "RepeatHeaderAndFirstColumn.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during workbook creation: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RepeatHeaderAndFirstColumnDemo.Run();
        }
    }
}
