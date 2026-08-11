// Title: Add a Thick Black Outline Border to a Header Row and Freeze It with Aspose.Cells for .NET
// Description: This example creates a new workbook, writes column titles to cells A1‑D1, defines the header range, applies a thick black outline using SetOutlineBorders, and then freezes the first row with FreezePanes so the styled header stays visible while scrolling. The file is saved as HeaderBorderAndFreeze.xlsx.
// Keywords: Aspose.Cells C# | SetOutlineBorders | FreezePanes | header row border | thick black border | freeze top row | Excel export .NET | range styling Aspose | worksheet freeze pane | border and freeze example
// Common Searches: how to add a border to a header row using Aspose.Cells | freeze first row in Excel with Aspose.Cells .NET | Aspose.Cells SetOutlineBorders example | FreezePanes usage in C# Aspose.Cells | apply thick border to range Aspose.Cells
// Developer Intent: Apply a visible border to the header range and lock that row so it remains on screen during scrolling.
// Use Cases: Financial reports where column titles need a distinct border and stay fixed while reviewing large data tables. | Dashboard exports that require the header to be highlighted and always visible for end‑users. | Printable worksheets where a frozen, bordered header ensures consistent labeling across pages.
// AI Prompts: Show C# code that adds a thick black outline to a header range and freezes the row with Aspose.Cells. | Generate an Aspose.Cells example that styles multiple header rows with different borders and freezes them. | Explain how to adjust the freeze pane position after applying borders to keep the border visible.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBorderAndFreeze
{
    // This example creates a new workbook, writes column titles to cells A1‑D1, defines the header range, applies a thick black outline using SetOutlineBorders, and then freezes the first row with FreezePanes so the styled header stays visible while scrolling. The file is saved as HeaderBorderAndFreeze.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some header data (first row)
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Quantity");
                cells["D1"].PutValue("Price");

                // Define the header range (A1:D1)
                AsposeRange headerRange = cells.CreateRange("A1:D1");

                // Apply a uniform thick black outline border around the header range
                headerRange.SetOutlineBorders(CellBorderType.Thick, Color.Black);

                // Freeze the first row so the header (with its border) stays visible while scrolling
                // Freeze at cell A2 (the first cell below the header), freezing 1 row and 0 columns
                worksheet.FreezePanes("A2", 1, 0);

                // Determine output file path
                string outputPath = "HeaderBorderAndFreeze.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
