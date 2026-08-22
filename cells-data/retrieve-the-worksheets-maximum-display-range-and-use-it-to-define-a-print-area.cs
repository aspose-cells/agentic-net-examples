// Title: Set the worksheet print area to its MaxDisplayRange using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that obtains a worksheet's MaxDisplayRange with Aspose.Cells, converts the range to A1 notation, and assigns it to PageSetup.PrintArea. | Show how to calculate the start and end cell addresses from MaxDisplayRange properties and apply them as the print area in a .NET Excel workbook.
// Common Searches: C# Aspose.Cells how to set print area based on worksheet's maximum display range | using MaxDisplayRange to define print area in Aspose.Cells .NET | Aspose.Cells retrieve MaxDisplayRange and set PageSetup.PrintArea programmatically | convert MaxDisplayRange indices to A1 address for printing with Aspose.Cells
// Tags: Aspose.Cells print area from MaxDisplayRange | C# retrieve worksheet MaxDisplayRange | PageSetup.PrintArea configuration Aspose.Cells | Excel A1 notation generation C# | define worksheet print region programmatically

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, adds sample data, obtains the worksheet's MaxDisplayRange, translates its bounds to A1-style addresses, sets PageSetup.PrintArea to that range, and saves the file.
class SetPrintAreaFromMaxDisplayRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data to create a display range
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["C5"].PutValue("Extra");

            // Retrieve the maximum display range (includes data, merged cells, shapes)
            AsposeRange maxRange = sheet.Cells.MaxDisplayRange;

            if (maxRange != null)
            {
                // Convert range indices to Excel cell addresses (e.g., A1, D10)
                string startAddress = CellsHelper.CellIndexToName(maxRange.FirstRow, maxRange.FirstColumn);
                string endAddress = CellsHelper.CellIndexToName(
                    maxRange.FirstRow + maxRange.RowCount - 1,
                    maxRange.FirstColumn + maxRange.ColumnCount - 1);

                // Define the print area using the calculated addresses
                sheet.PageSetup.PrintArea = $"{startAddress}:{endAddress}";
            }

            // Save the workbook with the defined print area
            workbook.Save("PrintAreaFromMaxDisplayRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
