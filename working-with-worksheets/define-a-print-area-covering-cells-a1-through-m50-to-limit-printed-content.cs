// Title: How to define a print area for cells A1:M50 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet print area to A1:M50 | limit printed range in Excel file with Aspose.Cells PageSetup.PrintArea | C# code to configure print area before saving workbook | example of setting print area for a specific range using Aspose.Cells .NET | configure worksheet print area in Aspose.Cells API
// Tags: Aspose.Cells PageSetup.PrintArea property | C# set Excel print area range | define worksheet print area Aspose.Cells | limit printed range Excel .NET | print area configuration Aspose.Cells

using System;
using Aspose.Cells;

namespace PrintAreaDemo
{
    // The sample creates a new workbook, fills cells with sample data, sets the worksheet's PageSetup.PrintArea to "A1:M50" to restrict the printed area, and saves the file as PrintAreaDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Optional: add some sample data to visualize the print area
            for (int row = 0; row < 60; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    // Fill cells with their address (e.g., "A1", "B1", ...)
                    string cellName = CellsHelper.CellIndexToName(row, col);
                    worksheet.Cells[cellName].PutValue(cellName);
                }
            }

            // Define the print area to cover cells A1 through M50
            // PageSetup.PrintArea expects a string in A1 notation
            worksheet.PageSetup.PrintArea = "A1:M50";

            // Optionally, set the print title rows/columns if needed
            // worksheet.PageSetup.PrintTitleRows = "1:1"; // repeat first row on each printed page
            // worksheet.PageSetup.PrintTitleColumns = "A:A"; // repeat first column on each printed page

            // Save the workbook to a file
            string outputPath = "PrintAreaDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("Print area set to A1:M50.");
        }
    }
}
