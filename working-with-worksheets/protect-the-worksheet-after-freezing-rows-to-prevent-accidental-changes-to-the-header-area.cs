// Title: Freeze Header Row and Protect Worksheet with Password using Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, freeze the first row as a header, apply full worksheet protection with a password, and save the file as ProtectedHeader.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells freeze row | protect worksheet password | freeze panes .NET | worksheet protection Aspose.Cells | C# Excel header freeze | Aspose.Cells example | Excel security Aspose
// Common Searches: Aspose.Cells freeze first row and protect | how to protect worksheet after freezing panes in C# | set password on Excel sheet with Aspose.Cells | freeze header and lock cells Aspose.Cells .NET | protect Excel workbook programmatically Aspose
// Developer Intent: Freeze the top row and secure the entire worksheet with a password.
// Use Cases: Create a read‑only report where the header stays visible while scrolling. | Distribute a template that locks all cells after freezing the header to avoid accidental edits. | Generate an export file for external partners with the header area frozen and password‑protected.
// AI Prompts: Provide C# code that freezes the first row and applies full worksheet protection with a password using Aspose.Cells. | Show how to freeze multiple header rows and protect only specific ranges in an Aspose.Cells workbook. | Explain how to programmatically unprotect a worksheet that was frozen and secured with a password in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a new workbook, freeze the first row as a header, apply full worksheet protection with a password, and save the file as ProtectedHeader.xlsx using Aspose.Cells for .NET.
public class ProtectHeaderDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row (header) – freeze panes at the second row (index 1)
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Protect the worksheet with all protection options and a password
            sheet.Protect(ProtectionType.All, "HeaderPass", null);

            // Save the workbook
            workbook.Save("ProtectedHeader.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ProtectHeaderDemo.Run();
    }
}
