// Title: Convert an existing XLSX workbook to ODS with Aspose.Cells for .NET and report the percentage size reduction
// AI Prompts: Load a .xlsx file using Aspose.Cells Workbook, save it as .ods, and print both file sizes. | Compute and display the percentage decrease in size after converting the workbook to ODS format with Aspose.Cells. | Write C# code that reads an Excel workbook, exports it to OpenDocument Spreadsheet, and shows the size reduction result.
// Common Searches: Aspose.Cells C# export Excel workbook to ODS and check file size | C# compare .xlsx and .ods file sizes after conversion using Aspose.Cells | measure storage savings when converting Excel to OpenDocument with Aspose.Cells .NET | sample code for saving workbook as ODS and getting size reduction percentage in C#
// Tags: Aspose.Cells save workbook as ODS | C# calculate file size reduction after format conversion | OpenDocument Spreadsheet export with Aspose.Cells | compare XLSX and ODS file sizes using .NET | measure storage savings with Aspose.Cells ODS output

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing XLSX workbook, saves it as an ODS file using Aspose.Cells, then retrieves and prints the original and converted file sizes along with the percentage reduction.
class Program
{
    static void Main()
    {
        // Paths for the original workbook and the ODS output
        string originalPath = "original.xlsx";
        string odsPath = "converted.ods";

        // Load the workbook that contains the converted range
        Workbook workbook = new Workbook(originalPath);

        // Save the workbook as ODS format
        workbook.Save(odsPath, SaveFormat.ODS);

        // Get file sizes
        long originalSize = new FileInfo(originalPath).Length;
        long odsSize = new FileInfo(odsPath).Length;

        // Calculate size reduction percentage
        double reductionPercent = 0;
        if (originalSize > 0)
        {
            reductionPercent = ((double)(originalSize - odsSize) / originalSize) * 100;
        }

        // Output the results
        Console.WriteLine($"Original file size: {originalSize} bytes");
        Console.WriteLine($"ODS file size: {odsSize} bytes");
        Console.WriteLine($"Size reduction: {reductionPercent:F2}%");
    }
}
