// Title: Copy a Range to a New Workbook and Set Password‑Protected Read‑Only Sheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a source workbook, copy cells A1:C3 to a new workbook, protect the destination worksheet with a password for read‑only access, and save the file as CopiedAndProtected.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range C# | protect worksheet password .NET | read‑only Excel sheet Aspose | copy cells between workbooks | Aspose.Cells range example | C# Excel protection Aspose
// Common Searches: Aspose.Cells copy range to another workbook | How to password‑protect a sheet with Aspose.Cells | C# code to create read‑only Excel file using Aspose | Copy and protect Excel range programmatically | Aspose.Cells example for sheet protection
// Developer Intent: Transfer a specific cell block into a fresh workbook and enforce password‑based read‑only protection on the target worksheet.
// Use Cases: Generate a client‑ready report by copying a data table from a template and locking the sheet to prevent edits. | Distribute chart source data in a separate file while safeguarding the original values with sheet protection. | Automate creation of secure Excel deliverables for external partners, copying only required ranges and applying a password.
// AI Prompts: Provide C# Aspose.Cells code that copies cells A1:C3 to a new workbook and protects the sheet with a password for read‑only access. | Show how to set different protection options after copying a range using Aspose.Cells for .NET. | Explain how to programmatically change the password or protection type on a worksheet created with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a source workbook, copy cells A1:C3 to a new workbook, protect the destination worksheet with a password for read‑only access, and save the file as CopiedAndProtected.xlsx using Aspose.Cells for .NET.
public class CopyRangeAndProtectSheet
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create the source workbook and populate a sample range (A1:C3)
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                sourceSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define the source range to copy (A1:C3)
        AsposeRange sourceRange = sourceSheet.Cells.CreateRange(0, 0, 3, 3);

        // Create the destination workbook
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the destination range (starting at A1)
        AsposeRange destinationRange = destinationSheet.Cells.CreateRange(0, 0, 3, 3);

        // Copy the source range into the destination range
        destinationRange.Copy(sourceRange);

        // Protect the destination worksheet with a password (read‑only access)
        string sheetPassword = "ReadOnly123";
        destinationSheet.Protect(ProtectionType.All, sheetPassword, null);

        // Save the new workbook
        string outputPath = "CopiedAndProtected.xlsx";
        destinationWorkbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
