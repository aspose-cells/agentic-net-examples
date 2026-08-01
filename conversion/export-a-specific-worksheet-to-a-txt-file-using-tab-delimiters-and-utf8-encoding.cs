// Title: C# – Export a Single Worksheet to a Tab‑Delimited UTF‑8 TXT File with Aspose.Cells
// Description: Shows how to pick a worksheet, copy it into a temporary workbook, configure TxtSaveOptions for a tab separator and UTF‑8 encoding, and write the sheet to a .txt file.
// Keywords: Aspose.Cells | C# export worksheet to txt | tab delimited text | UTF-8 encoding | TxtSaveOptions | single sheet export | Excel to txt conversion | .NET Aspose.Cells example | save worksheet as text file
// Common Searches: Aspose.Cells export specific sheet to txt | C# save worksheet as tab separated values | TxtSaveOptions tab delimiter example | how to export only one worksheet to UTF‑8 text file | Aspose.Cells .NET export active sheet to txt
// Developer Intent: Save a selected worksheet as a UTF‑8 encoded, tab‑separated text file.
// Use Cases: Create a plain‑text report of product data from a particular sheet for downstream analytics. | Generate a tab‑separated feed for a legacy system that accepts only UTF‑8 text files. | Extract a single sheet from a large workbook to share with non‑Excel users.
// AI Prompts: Write C# code using Aspose.Cells that exports the active worksheet to a .txt file with tab delimiters and UTF‑8 encoding, ensuring only that sheet is saved. | Provide an Aspose.Cells for .NET example that copies a specific worksheet into a new workbook and saves it as a tab‑delimited UTF‑8 text file. | Explain how to set TxtSaveOptions for tab separation and UTF‑8 encoding when converting a worksheet to a TXT file.

using System;
using System.Text;
using Aspose.Cells;
using System.IO;

// Shows how to pick a worksheet, copy it into a temporary workbook, configure TxtSaveOptions for a tab separator and UTF‑8 encoding, and write the sheet to a .txt file.
class ExportWorksheetToTxt
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();

            // First worksheet (will not be exported)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Name");
            sheet1.Cells["B1"].PutValue("Age");
            sheet1.Cells["A2"].PutValue("John");
            sheet1.Cells["B2"].PutValue(30);

            // Second worksheet (the one we want to export)
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Product");
            sheet2.Cells["B1"].PutValue("Price");
            sheet2.Cells["A2"].PutValue("Laptop");
            sheet2.Cells["B2"].PutValue(999.99);

            // Set the second worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = sheet2.Index;

            // Create a temporary workbook containing only the active worksheet
            Workbook exportWb = new Workbook();
            exportWb.Worksheets.Clear();                     // remove default sheet
            exportWb.Worksheets.AddCopy(sheet2.Name);        // copy the active sheet by name

            // Configure TxtSaveOptions: tab delimiter, UTF‑8 encoding
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = '\t',               // use tab as delimiter
                Encoding = Encoding.UTF8        // UTF‑8 encoding
            };

            // Define output path and ensure the directory exists
            string outputPath = "SecondSheet.txt";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the active worksheet to a TXT file
            exportWb.Save(outputPath, saveOptions);
            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
