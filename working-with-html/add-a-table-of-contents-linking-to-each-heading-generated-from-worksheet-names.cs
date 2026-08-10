// Title: Aspose.Cells for .NET – Create a Table of Contents worksheet with hyperlinks to every sheet
// Description: C# example that builds a new Workbook, adds sample worksheets, inserts a "Table of Contents" sheet at the first position, lists each sheet name in column A and creates a hyperlink to the sheet's A1 cell, then saves the file as WorkbookWithTOC.xlsx.
// Keywords: Aspose.Cells .NET | C# Excel Table of Contents | worksheet hyperlinks | generate Excel TOC programmatically | navigation sheet Aspose.Cells | Excel workbook automation | hyperlink to sheet A1 | insert worksheet at index 0 | GitHub Aspose.Cells example | coding‑agent Excel TOC
// Common Searches: how to add a table of contents sheet with links using Aspose.Cells | Aspose.Cells create TOC for multi‑sheet workbook C# | C# generate Excel navigation sheet with hyperlinks | insert Table of Contents as first worksheet Aspose.Cells | programmatic Excel TOC example .NET
// Developer Intent: Programmatically add a clickable Table of Contents sheet to an Excel workbook created with Aspose.Cells.
// Use Cases: Produce a navigable TOC for large reports so readers can jump directly to each data section. | Build a reusable template that automatically inserts a summary sheet with links for any generated workbook. | Create an internal dashboard where the first sheet lists all analysis tabs with one‑click access.
// AI Prompts: Show how to add sheet index numbers and custom styling (font size, color) to each TOC entry. | Give an example that skips hidden or template worksheets when building the Table of Contents. | Explain how to add a second column with user‑defined descriptions while keeping the hyperlinks functional.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that builds a new Workbook, adds sample worksheets, inserts a "Table of Contents" sheet at the first position, lists each sheet name in column A and creates a hyperlink to the sheet's A1 cell, then saves the file as WorkbookWithTOC.xlsx.
    public class TableOfContentsDemo
    {
        public static void Main()
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

            // Access the worksheet collection
            WorksheetCollection sheets = workbook.Worksheets;

            // Add sample worksheets that will become headings in the TOC
            sheets.Add("SalesData");
            sheets.Add("Inventory");
            sheets.Add("Summary Report");

            // Insert a new worksheet at the beginning to serve as the Table of Contents
            Worksheet tocSheet = sheets.Insert(0, SheetType.Worksheet, "TableOfContents");

            // Set a title for the TOC sheet
            Cell titleCell = tocSheet.Cells["A1"];
            titleCell.PutValue("Table of Contents");
            // Apply bold font to the title
            Style titleStyle = titleCell.GetStyle();
            titleStyle.Font.IsBold = true;
            titleCell.SetStyle(titleStyle);

            // Start listing entries from row 3 (index 2) to leave space after the title
            int tocRow = 2;

            // Iterate over all worksheets except the TOC sheet itself
            for (int i = 1; i < sheets.Count; i++)
            {
                Worksheet sheet = sheets[i];

                // Write the sheet name in column A
                tocSheet.Cells[tocRow, 0].PutValue(sheet.Name);

                // Create a hyperlink from the cell to the corresponding sheet's A1 cell
                string address = $"'{sheet.Name}'!A1";
                tocSheet.Hyperlinks.Add(tocRow, 0, 1, 1, address);

                tocRow++;
            }

            // Save the workbook
            string outputPath = "WorkbookWithTOC.xlsx";
            workbook.Save(outputPath);
        }
    }
}
