// Title: C# – Remove Worksheet Printer Settings with Aspose.Cells for .NET
// Description: This Aspose.Cells for .NET example creates a workbook, accesses the first worksheet, optionally adds data, and clears the embedded printer configuration by setting worksheet.PageSetup.PrinterSettings to null before saving the file.
// Keywords: Aspose.Cells remove printer settings | C# clear worksheet printer configuration | PageSetup.PrinterSettings null | delete embedded printer settings Excel | Aspose.Cells .NET printer setup | Excel workbook printer reset | remove printer config Aspose.Cells
// Common Searches: how to clear printer settings in Aspose.Cells C# | set PageSetup.PrinterSettings to null example | remove embedded printer configuration from Excel using Aspose | Aspose.Cells delete worksheet printer settings | C# code to reset printer settings in Excel workbook
// Developer Intent: Clear the worksheet's embedded printer configuration by assigning null to its PageSetup.PrinterSettings property.
// Use Cases: Distribute a workbook without any printer-specific data attached. | Reset printer settings after copying worksheets between workbooks to avoid inheriting source printer information. | Prepare template files for end users so the Excel file opens with neutral printer defaults.
// AI Prompts: Generate C# code that uses Aspose.Cells to remove printer settings from a specific worksheet. | Explain why assigning null to worksheet.PageSetup.PrinterSettings eliminates embedded printer configuration in an Excel file. | List steps to verify that printer settings have been cleared after saving a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This Aspose.Cells for .NET example creates a workbook, accesses the first worksheet, optionally adds data, and clears the embedded printer configuration by setting worksheet.PageSetup.PrinterSettings to null before saving the file.
public class RemovePrinterSettingsDemo
{
    public static void Main(string[] args)
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
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data (optional)
        worksheet.Cells["A1"].PutValue("Demo for removing printer settings");

        // Remove the embedded printer configuration by setting it to null
        worksheet.PageSetup.PrinterSettings = null;

        // Define output file path
        string outputPath = "RemovedPrinterSettings.xlsx";

        // Save the workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
    }
}
