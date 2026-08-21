// Title: Aspose.Cells for .NET – Set Worksheet Custom Paper Size to 500 pt × 700 pt (C#)
// Description: Learn how to define a custom worksheet page size of 500 points width and 700 points height using Aspose.Cells for .NET. The example converts points to inches, applies PageSetup.CustomPaperSize, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells custom paper size | PageSetup.CustomPaperSize C# | set worksheet page dimensions .NET | 500 pt 700 pt Excel | convert points to inches Aspose | C# Excel custom page size
// Common Searches: Aspose.Cells set custom paper size 500x700 points | PageSetup.CustomPaperSize example C# | how to convert points to inches in Aspose.Cells | C# define worksheet page size Aspose | save workbook after custom page size Aspose.Cells
// Developer Intent: Create a worksheet with a 500 pt × 700 pt custom paper size and save the workbook.
// Use Cases: Printing reports on non‑standard paper formats | Generating Excel files for custom‑sized marketing collateral | Automating batch creation of spreadsheets that must match specific page dimensions
// AI Prompts: Show a C# snippet that sets a worksheet's custom paper size to 500 points width and 700 points height using Aspose.Cells. | Provide an example that adds sample data, applies the custom page size, and exports the workbook to PDF while preserving dimensions. | Explain how to read and modify an existing worksheet's custom paper size programmatically with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Learn how to define a custom worksheet page size of 500 points width and 700 points height using Aspose.Cells for .NET. The example converts points to inches, applies PageSetup.CustomPaperSize, and saves the workbook as an Excel file.
class CustomPaperSizeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Convert points to inches (1 point = 1/72 inch)
        double widthInInches = 500.0 / 72.0;   // ≈ 6.94444 inches
        double heightInInches = 700.0 / 72.0;  // ≈ 9.72222 inches

        // Set the custom paper size for the worksheet
        worksheet.PageSetup.CustomPaperSize(widthInInches, heightInInches);

        // Save the workbook
        workbook.Save("CustomPaperSize.xlsx");
    }
}
