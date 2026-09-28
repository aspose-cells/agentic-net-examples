// Title: Enable printed gridlines and set high‑resolution (600 DPI) print quality for an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that turns on gridlines for printed pages and configures the worksheet to use 600 DPI print quality before saving. | Show how to modify a worksheet's PageSetup in Aspose.Cells so that the printed output includes gridlines and is rendered at high resolution.
// Common Searches: Aspose.Cells C# enable gridlines when printing Excel worksheet | Set print DPI to 600 using Aspose.Cells .NET | Configure PageSetup for high‑resolution printing in Aspose.Cells | Print Excel file with gridlines and high quality using Aspose.Cells C# example | Aspose.Cells page setup print quality 600 DPI code sample
// Tags: Aspose.Cells worksheet print gridlines | Aspose.Cells set print DPI 600 | Aspose.Cells PageSetup high resolution printing | C# Aspose.Cells print quality configuration | Excel workbook page setup Aspose.Cells

using Aspose.Cells;
using System;

// // Loads input.xlsx, enables gridlines in printed pages, sets print quality to 600 DPI via PageSetup, and saves the modified workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable gridlines in the printed output
        sheet.PageSetup.PrintGridlines = true;

        // Set high print quality (e.g., 600 DPI)
        sheet.PageSetup.PrintQuality = 600;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
