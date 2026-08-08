// Title: Add DetailLink Hyperlinks from Master to Detail Worksheets with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a workbook with Master and Detail sheets, populate them with sample data, and use Aspose.Cells Hyperlinks.Add to create DetailLink navigation links from each master row to its corresponding detail row, then save the file.
// Keywords: Aspose.Cells | DetailLink | Hyperlink | Master sheet | Detail sheet | C# | ASP.NET | Excel navigation | Hyperlinks.Add example
// Common Searches: Aspose.Cells add hyperlink between worksheets C# | DetailLink master detail example .NET | How to create navigation links in Excel using Aspose.Cells | C# Aspose.Cells Hyperlinks.Add usage
// Developer Intent: Create DetailLink navigation hyperlinks from master rows to matching detail rows in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Build an interactive report where clicking an ID in the Master sheet jumps to the related record in the Detail sheet. | Automate generation of Excel dashboards that provide quick access to underlying transaction data. | Enable end‑users to navigate large exported workbooks (e.g., ERP or financial statements) without manual searching.
// AI Prompts: Generate C# code that loops through all master rows and adds DetailLink hyperlinks to the matching detail rows using Aspose.Cells. | Explain how to set custom display text for a hyperlink while keeping the underlying DetailLink address in Aspose.Cells. | Show an example of using Aspose.Cells Hyperlinks.Add to link a master sheet cell to a specific cell on a detail worksheet.

using System;
using Aspose.Cells;

// Demonstrates how to generate a workbook with Master and Detail sheets, populate them with sample data, and use Aspose.Cells Hyperlinks.Add to create DetailLink navigation links from each master row to its corresponding detail row, then save the file.
class DetailLinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------- Master Worksheet --------------------
        Worksheet masterSheet = workbook.Worksheets[0];
        masterSheet.Name = "Master";

        // Header
        masterSheet.Cells["A1"].PutValue("ID");
        masterSheet.Cells["B1"].PutValue("Name");

        // Sample master rows
        masterSheet.Cells["A2"].PutValue(1);
        masterSheet.Cells["B2"].PutValue("Item 1");
        masterSheet.Cells["A3"].PutValue(2);
        masterSheet.Cells["B3"].PutValue("Item 2");

        // -------------------- Detail Worksheet --------------------
        Worksheet detailSheet = workbook.Worksheets.Add("Detail");

        // Header
        detailSheet.Cells["A1"].PutValue("MasterID");
        detailSheet.Cells["B1"].PutValue("DetailInfo");

        // Sample detail rows corresponding to master IDs
        detailSheet.Cells["A2"].PutValue(1);
        detailSheet.Cells["B2"].PutValue("Detail for Item 1");
        detailSheet.Cells["A3"].PutValue(2);
        detailSheet.Cells["B3"].PutValue("Detail for Item 2");

        // -------------------- Create Hyperlinks (DetailLink) --------------------
        // Hyperlink address format: 'SheetName'!CellReference
        // Link master row 2 to detail row 2
        masterSheet.Hyperlinks.Add("A2", 1, 1, "'Detail'!A2");
        // Link master row 3 to detail row 3
        masterSheet.Hyperlinks.Add("A3", 1, 1, "'Detail'!A3");

        // Optional: change displayed text for clarity
        masterSheet.Cells["A2"].PutValue("Go to Detail 1");
        masterSheet.Cells["A3"].PutValue("Go to Detail 2");

        // Save the workbook
        workbook.Save("DetailLinkDemo.xlsx");
    }
}
