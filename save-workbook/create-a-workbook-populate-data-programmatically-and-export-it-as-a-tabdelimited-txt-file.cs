// Title: Export a single worksheet to a UTF-8 tab-delimited TXT file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, writes product data into cells, and saves only the active sheet as a UTF-8 tab-separated text file using Aspose.Cells. | Show how to configure TxtSaveOptions by setting Separator to a tab character, specifying Encoding.UTF8, and disabling ExportAllSheets to export a single sheet to .txt.
// Common Searches: Aspose.Cells C# export active worksheet to tab separated values file | How to save a workbook as a .txt file with tab delimiter using Aspose.Cells .NET | Set UTF-8 encoding for TxtSaveOptions when exporting to text in Aspose.Cells | Export only the first sheet to a tab-delimited text file with Aspose.Cells | C# example for creating a workbook and exporting to TSV using Aspose.Cells
// Tags: export worksheet to tab-delimited txt Aspose.Cells | TxtSaveOptions separator tab C# | UTF-8 text export Aspose.Cells | save active sheet as txt Aspose.Cells | populate workbook cells programmatically C#

using System;
using Aspose.Cells;
using System.Text;

namespace AsposeCellsTxtExportExample
{
    // The sample creates a new Workbook, fills cells A1‑B4 with product names and prices, configures TxtSaveOptions to use a tab separator, UTF-8 encoding, and to export only the active sheet, then saves the result as ExportedData.txt.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(599.99);
            worksheet.Cells["A4"].PutValue("Tablet");
            worksheet.Cells["B4"].PutValue(399.99);

            // Configure text save options for tab‑delimited export
            TxtSaveOptions txtOptions = new TxtSaveOptions();
            txtOptions.Separator = '\t';               // Use tab as the delimiter
            txtOptions.Encoding = Encoding.UTF8;       // Set desired encoding
            txtOptions.ExportAllSheets = false;        // Export only the active sheet

            // Save the workbook as a tab‑delimited TXT file
            string outputPath = "ExportedData.txt";
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
    }
}
