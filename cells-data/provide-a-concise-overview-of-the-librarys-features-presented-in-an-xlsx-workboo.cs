using System;
using Aspose.Cells;

class FeaturesOverview
{
    static void Main()
    {
        // Create a new workbook instance (default format is Xlsx)
        Workbook wb = new Workbook();

        // Access the default worksheet and give it a meaningful name
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Features";

        // Write a title and column headers
        ws.Cells["A1"].PutValue("Aspose.Cells for .NET - Feature Overview");
        ws.Cells["A2"].PutValue("Feature");
        ws.Cells["B2"].PutValue("Description");

        // Populate a concise list of key library capabilities
        int row = 3;

        ws.Cells[row, 0].PutValue("Workbook Creation");
        ws.Cells[row, 1].PutValue("Create, load, and save Excel files in various formats.");
        row++;

        ws.Cells[row, 0].PutValue("Worksheets Management");
        ws.Cells[row, 1].PutValue("Add, remove, rename, and access worksheets.");
        row++;

        ws.Cells[row, 0].PutValue("Cell Operations");
        ws.Cells[row, 1].PutValue("Read/write values, formulas, styles, and formatting.");
        row++;

        ws.Cells[row, 0].PutValue("Styling");
        ws.Cells[row, 1].PutValue("Create and apply styles, themes, and custom colors.");
        row++;

        ws.Cells[row, 0].PutValue("Formulas");
        ws.Cells[row, 1].PutValue("Calculate, parse, and manage formulas, including array formulas.");
        row++;

        ws.Cells[row, 0].PutValue("Data Import/Export");
        ws.Cells[row, 1].PutValue("Import/Export XML, JSON, CSV, and integrate with external data sources.");
        row++;

        ws.Cells[row, 0].PutValue("Charts");
        ws.Cells[row, 1].PutValue("Create and customize various chart types.");
        row++;

        ws.Cells[row, 0].PutValue("PivotTables");
        ws.Cells[row, 1].PutValue("Create, configure, and refresh PivotTables.");
        row++;

        ws.Cells[row, 0].PutValue("Protection");
        ws.Cells[row, 1].PutValue("Protect workbook, worksheets, and cells with passwords.");
        row++;

        ws.Cells[row, 0].PutValue("Macros");
        ws.Cells[row, 1].PutValue("Add, remove, and manage VBA macros.");
        row++;

        ws.Cells[row, 0].PutValue("Digital Signatures");
        ws.Cells[row, 1].PutValue("Add and verify digital signatures.");
        row++;

        ws.Cells[row, 0].PutValue("AI Integration");
        ws.Cells[row, 1].PutValue("Use CellsAI for summarization and Q&A on spreadsheets.");
        row++;

        ws.Cells[row, 0].PutValue("Export Options");
        ws.Cells[row, 1].PutValue("Save to PDF, HTML, EPUB, images, and more with fine‑grained options.");
        row++;

        // Adjust column widths for readability
        ws.AutoFitColumns();

        // Save the workbook to disk using the provided Save method
        wb.Save("FeaturesOverview.xlsx");

        // Release resources
        wb.Dispose();
    }
}