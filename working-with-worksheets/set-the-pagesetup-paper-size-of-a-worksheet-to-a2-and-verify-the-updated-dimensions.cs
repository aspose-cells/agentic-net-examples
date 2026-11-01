// Title: How to set a worksheet's page setup to A2 paper size and read its dimensions using Aspose.Cells for .NET
// AI Prompts: Set the PageSetup.PaperSize of a worksheet to PaperA2 and print the resulting PaperWidth and PaperHeight in points. | Check whether the worksheet's PaperSize property equals PaperA2 and output a boolean confirmation. | Create a new workbook, change the first sheet's page setup to A2, and display the actual page dimensions for verification.
// Common Searches: Aspose.Cells .NET set worksheet page size to A2 and get width height in points | How to verify paper size A2 after configuring PageSetup in Aspose.Cells | Retrieve paper dimensions of an Excel sheet after setting A2 paper size with Aspose.Cells | C# code to check worksheet PaperSize property equals PaperA2 in Aspose.Cells | Aspose.Cells page setup paper width and height values for A2 size
// Tags: Aspose.Cells set worksheet paper size A2 | Aspose.Cells get page dimensions in points | verify worksheet PaperSize property .NET | configure PageSetup for A2 size Excel | C# Aspose.Cells page setup API usage

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PaperSize to PaperA2, confirms the setting with a boolean check, and prints the worksheet's PaperWidth and PaperHeight values (in points) for verification.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page‑setup paper size to A2
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA2;

        // Verify that the paper size was set to A2
        bool isA2 = sheet.PageSetup.PaperSize == PaperSizeType.PaperA2;
        Console.WriteLine("Paper size set to A2: " + isA2);

        // Output the actual dimensions (in points) for additional verification
        double widthPoints = sheet.PageSetup.PaperWidth;
        double heightPoints = sheet.PageSetup.PaperHeight;
        Console.WriteLine($"Paper width (points): {widthPoints}");
        Console.WriteLine($"Paper height (points): {heightPoints}");
    }
}
