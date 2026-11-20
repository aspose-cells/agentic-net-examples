// Title: Generate a QuotePrefix report in C# with Aspose.Cells: list sheet name, row and column indices of cells that have Style.QuotePrefix enabled
// Common Searches: C# Aspose.Cells how to find cells with QuotePrefix flag and export their addresses | list all cells where Style.QuotePrefix is true using Aspose.Cells .NET | create a worksheet report of quoted prefix cells with row and column numbers in Aspose.Cells | Aspose.Cells iterate workbook cells and detect QuotePrefix property in C# | extract zero‑based row and column indices of cells that have quote prefix in Excel via Aspose.Cells
// Tags: detect QuotePrefix cells Aspose.Cells | generate quoted‑prefix report C# | iterate workbook used range Aspose.Cells | export cell coordinates Aspose.Cells | create summary worksheet Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads (or creates) an Excel workbook, adds a new worksheet named "QuotePrefixReport", writes a header row, and then scans every worksheet's used range. For each cell whose Style.QuotePrefix property is true, it records the sheet name, zero‑based row index, column index, and cell address on the report sheet. Finally, the workbook is saved with the generated report.
class QuotePrefixReport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // Load without password; if the file is password‑protected this will throw,
                // which is caught by the outer try‑catch.
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                // Add some sample data (optional)
                ws.Cells["A1"].PutValue("Sample");
            }

            // Add a new worksheet that will hold the report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "QuotePrefixReport";

            // Write header row for the report
            reportSheet.Cells[0, 0].PutValue("Sheet");
            reportSheet.Cells[0, 1].PutValue("RowIndex");
            reportSheet.Cells[0, 2].PutValue("ColumnIndex");
            reportSheet.Cells[0, 3].PutValue("CellAddress");

            int reportRow = 1; // start writing data after header

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the report sheet itself to avoid self‑reference
                if (sheet.Name == "QuotePrefixReport")
                    continue;

                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check if the cell's style has QuotePrefix set to true
                        if (cell.GetStyle().QuotePrefix)
                        {
                            // Record the cell's location in the report sheet
                            reportSheet.Cells[reportRow, 0].PutValue(sheet.Name);
                            reportSheet.Cells[reportRow, 1].PutValue(row);
                            reportSheet.Cells[reportRow, 2].PutValue(col);
                            reportSheet.Cells[reportRow, 3].PutValue(cell.Name); // e.g., "A1"
                            reportRow++;
                        }
                    }
                }
            }

            // Save the workbook, now containing the report
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
