// Title: Create a new workbook, define a worksheet‑scoped named range (A1:C5), and save it as an XLSX file with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that initializes a workbook, creates a range A1:C5 on the first worksheet, assigns it a worksheet‑only name, and writes the file in XLSX format. | Show how to add a worksheet‑level named range in a fresh workbook and persist the workbook as an .xlsx file using the Aspose.Cells API.
// Common Searches: Aspose.Cells how to add a worksheet scoped named range in C# | C# Aspose.Cells save new workbook with named range to xlsx | define A1:C5 range with worksheet level scope using Aspose.Cells .NET | create workbook and set worksheet‑only named range Aspose.Cells example | Aspose.Cells .NET export workbook after naming a range on a single sheet
// Tags: worksheet‑level named range Aspose.Cells | new workbook generation Aspose.Cells | export workbook to XLSX Aspose.Cells | range A1:C5 definition Aspose.Cells | named range per worksheet Aspose.Cells

using System;
using Aspose.Cells;

// // This program creates a new workbook, defines a worksheet‑scoped named range covering cells A1:C5, assigns it the name "MyWorksheetRange", and saves the workbook as "WorksheetScopedNamedRange.xlsx" in XLSX format.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (default worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range (e.g., A1:C5)
            int firstRow = 0;          // Row index starts at 0 (A1)
            int firstColumn = 0;       // Column index starts at 0 (A)
            int totalRows = 5;         // Number of rows in the range
            int totalColumns = 3;      // Number of columns in the range

            // Create a worksheet‑scoped named range
            Aspose.Cells.Range namedRange = worksheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);
            namedRange.Name = "MyWorksheetRange"; // This name is scoped to the worksheet

            // Save the workbook in XLSX format
            workbook.Save("WorksheetScopedNamedRange.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
