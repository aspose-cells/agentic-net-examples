// Title: Create a non‑contiguous union range (A1:B2, D4:E5) and apply a bold font style using Aspose.Cells for .NET
// Common Searches: Aspose.Cells create union range from multiple addresses C# | apply bold font to noncontiguous cells using Aspose.Cells .NET | Worksheet.Cells.CreateRange with comma‑separated addresses example | style cells in a union range Aspose.Cells C# tutorial | set bold style for cells A1:B2 and D4:E5 using Aspose.Cells
// Tags: union range creation with Worksheet.Cells.CreateRange | bold font styling for Excel cells using Aspose.Cells | non‑contiguous range formatting in .NET | apply style to multiple address blocks Aspose.Cells | Excel workbook save after range styling Aspose

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a union range covering A1:B2 and D4:E5 with Worksheet.Cells.CreateRange, builds a bold font style, applies it to each cell in the union range, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a union range that combines the two address blocks
            // Aspose.Cells allows multiple areas separated by commas in the address string
            Aspose.Cells.Range unionRange = sheet.Cells.CreateRange("A1:B2,D4:E5");

            // Define a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Apply the bold style to every cell in the union range
            foreach (Cell cell in unionRange)
            {
                cell.SetStyle(boldStyle);
            }

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
