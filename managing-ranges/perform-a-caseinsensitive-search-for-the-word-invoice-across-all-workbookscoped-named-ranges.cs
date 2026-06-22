using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Invoice #123");
            sheet.Cells["A2"].PutValue("No invoice here");
            sheet.Cells["B1"].PutValue("Other");
            sheet.Cells["B2"].PutValue("invoice details");
            sheet.Cells["A3"].PutValue("Invoice summary");
            sheet.Cells["A4"].PutValue("Summary");

            // Create global (workbook‑scoped) named range "GlobalRange1" covering A1:B2
            int idx1 = workbook.Worksheets.Names.Add("GlobalRange1");
            Name globalName1 = workbook.Worksheets.Names[idx1];
            globalName1.RefersTo = "=Sheet1!$A$1:$B$2";
            globalName1.SheetIndex = 0; // 0 = workbook scope

            // Create another global named range "GlobalRange2" covering A3:A4
            int idx2 = workbook.Worksheets.Names.Add("GlobalRange2");
            Name globalName2 = workbook.Worksheets.Names[idx2];
            globalName2.RefersTo = "=Sheet1!$A$3:$A$4";
            globalName2.SheetIndex = 0;

            // Iterate over all defined names and process only workbook‑scoped ones
            foreach (Name name in workbook.Worksheets.Names)
            {
                if (name.SheetIndex != 0) // skip non‑global names
                    continue;

                // Retrieve the range that the name refers to
                AsposeRange range = name.GetRange();
                if (range == null)
                    continue;

                // Configure FindOptions for a case‑insensitive, contains search
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.Contains,
                    CaseSensitive = false
                };

                // Restrict the search to the current named range
                CellArea searchArea = new CellArea
                {
                    StartRow = range.FirstRow,
                    StartColumn = range.FirstColumn,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };
                findOptions.SetRange(searchArea);

                // Perform the search for the word "invoice"
                Cell foundCell = sheet.Cells.Find("invoice", null, findOptions);
                while (foundCell != null)
                {
                    Console.WriteLine($"Found in named range '{name.Text}' at cell {foundCell.Name}");
                    // Continue searching after the current cell
                    foundCell = sheet.Cells.Find("invoice", foundCell, findOptions);
                }
            }

            // Save the workbook (lifecycle: save)
            string outputPath = "InvoiceSearchResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}