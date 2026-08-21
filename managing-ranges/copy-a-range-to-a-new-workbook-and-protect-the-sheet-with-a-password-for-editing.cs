// Title: Copy a Range to a New Workbook and Password‑Protect the Sheet with Aspose.Cells for .NET (C#)
// Description: Creates a source workbook, fills a 3×3 range (A1:C3), copies it to a destination range (E5:G7) in a new workbook, applies full protection with the password "myPassword", and saves the file as CopiedAndProtected.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range C# | Aspose.Cells protect worksheet password | copy cells to another workbook | C# Aspose.Cells range.Copy example | sheet protection Aspose.Cells .NET | Aspose.Cells copy and lock sheet | Aspose.Cells workbook cloning | Aspose.Cells range protection options
// Common Searches: Aspose.Cells copy range to another workbook C# | How to protect an Aspose.Cells worksheet with a password | Copy cells and set sheet protection using Aspose.Cells .NET | C# copy range and lock sheet Aspose.Cells | Aspose.Cells copy range preserving formatting
// Developer Intent: Copy a defined cell range from a source workbook into a new workbook at a specific address, then apply full sheet protection with a password.
// Use Cases: Generate a report by copying a template block into a fresh workbook and locking the sheet to prevent edits. | Create a read‑only snapshot of selected data for distribution, ensuring the copied range cannot be modified. | Automate data extraction where a range is moved to a separate file and the worksheet is secured to maintain data integrity.
// AI Prompts: Write C# code with Aspose.Cells that copies range A1:D10 from a source workbook to range H5:K14 in a new workbook and protects the destination sheet with password 'Secure123'. | Show how to copy multiple non‑contiguous ranges into a new workbook and apply worksheet protection using Aspose.Cells for .NET. | Explain how to copy a range, keep its formatting, and set selective protection options (e.g., allow sorting) with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a source workbook, fills a 3×3 range (A1:C3), copies it to a destination range (E5:G7) in a new workbook, applies full protection with the password "myPassword", and saves the file as CopiedAndProtected.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create source workbook and populate a range with sample data
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    srcSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the source range (A1:C3)
            AsposeRange srcRange = srcSheet.Cells.CreateRange(0, 0, 3, 3);

            // Create destination workbook
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];

            // Define the destination range (E5:G7) where the data will be copied
            AsposeRange destRange = destSheet.Cells.CreateRange(4, 4, 3, 3);

            // Copy the source range to the destination range
            srcRange.Copy(destRange);

            // Protect the destination worksheet with a password for editing
            destSheet.Protect(ProtectionType.All, "myPassword", null);

            // Save the resulting workbook
            destWb.Save("CopiedAndProtected.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
