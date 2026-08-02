// Title: C# – Enable Print Gridlines & High‑Resolution (300 DPI) Printing with Aspose.Cells .NET
// Description: Demonstrates how to turn on PrintGridlines and set PrintQuality to 300 DPI via the Worksheet.PageSetup object, add optional content, and save the workbook as GridlinesHighQuality.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | .NET Excel printing | PrintGridlines | PrintQuality 300 DPI | high resolution Excel print | Worksheet PageSetup | gridlines in printed Excel | Excel to PDF high DPI | Aspose.Cells print settings | C# Excel export
// Common Searches: Aspose.Cells enable printed gridlines C# | set 300 DPI print quality Aspose.Cells .NET | how to configure page setup print settings Aspose.Cells | C# code for high‑resolution Excel printing with Aspose.Cells | print gridlines Excel workbook using Aspose.Cells
// Developer Intent: Turn on gridlines for printed output and configure the worksheet to print at high DPI resolution.
// Use Cases: Produce audit reports where visible gridlines aid data verification. | Create marketing PDFs from Excel sheets that require crisp, high‑resolution graphics. | Prepare spreadsheets for professional printing, ensuring both gridlines and DPI meet publishing standards.
// AI Prompts: Generate C# code that toggles PrintGridlines and sets PrintQuality to 600 DPI with Aspose.Cells. | Write a method to adjust PrintQuality based on a user‑selected quality level (e.g., 150, 300, 600 DPI). | Explain the impact of PrintQuality on file size and visual clarity when exporting to PDF using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to turn on PrintGridlines and set PrintQuality to 300 DPI via the Worksheet.PageSetup object, add optional content, and save the workbook as GridlinesHighQuality.xlsx using Aspose.Cells for .NET.
public class EnableGridlinesAndPrintQuality
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable gridlines in the printed output
        worksheet.PageSetup.PrintGridlines = true;

        // Set a high print quality (e.g., 300 DPI)
        worksheet.PageSetup.PrintQuality = 300;

        // (Optional) Add sample data to visualize the gridlines
        worksheet.Cells["A1"].PutValue("Gridlines and high‑quality print demo");

        // Save the workbook
        string outputPath = "GridlinesHighQuality.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
