// Title: How to print both gridlines and row/column headings in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, sets PageSetup.PrintGridlines and PageSetup.PrintHeadings to true, and saves the file. | Show how to configure Aspose.Cells Worksheet.PageSetup to enable printing of gridlines and headings together before exporting the workbook.
// Common Searches: Aspose.Cells C# enable printing of gridlines and row headings in the same printout | set PrintHeadings and PrintGridlines together using Aspose.Cells .NET | how to replicate on-screen view when printing Excel with Aspose.Cells | C# Aspose.Cells PageSetup print options for gridlines and headings | save workbook after turning on PrintGridlines and PrintHeadings in Aspose.Cells
// Tags: Aspose.Cells PageSetup PrintGridlines | Aspose.Cells PageSetup PrintHeadings | C# print Excel gridlines with Aspose.Cells | C# print Excel headings with Aspose.Cells | match on-screen view when printing Excel using Aspose.Cells

using Aspose.Cells;

// Loads an existing workbook, accesses the first worksheet, enables printing of both gridlines and row/column headings via the worksheet's PageSetup, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        var worksheet = workbook.Worksheets[0];

        // Enable printing of gridlines
        worksheet.PageSetup.PrintGridlines = true;

        // Enable printing of row and column headings
        worksheet.PageSetup.PrintHeadings = true;

        // Save the workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
