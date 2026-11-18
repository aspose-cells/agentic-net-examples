// Title: Combine multiple Excel workbooks into a single PDF while keeping worksheet order using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge several .xlsx files into one PDF preserving sheet sequence | How to copy worksheets from multiple workbooks into a master workbook and export to PDF in .NET | C# code to combine Excel worksheets from different files into a single PDF with Aspose.Cells | Preserve original worksheet order when converting multiple Excel workbooks to PDF using Aspose.Cells | Create merged PDF from multiple Excel files in C# without losing sheet order
// Tags: addcopy worksheet Aspose.Cells | merge workbooks into single PDF | export combined workbook to PDF | master workbook creation Aspose.Cells | combine .xlsx files into PDF C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a list of Excel file paths, creates an empty master workbook, removes its default sheet, and iteratively copies each worksheet from the source workbooks using AddCopy to retain the original sheet order. After ensuring the output directory exists, the combined workbook is saved as a single PDF. The code handles missing source files and runtime exceptions.
class ExcelToPdfMerger
{
    static void Main()
    {
        try
        {
            // Paths of the Excel files to be merged (order matters)
            string[] excelFiles = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Create a new empty workbook that will hold all worksheets
            Workbook masterWorkbook = new Workbook();

            // Remove the default empty sheet created with a new workbook
            if (masterWorkbook.Worksheets.Count > 0)
                masterWorkbook.Worksheets.RemoveAt(0);

            // Iterate through each source Excel file in the specified order
            foreach (string filePath in excelFiles)
            {
                // Verify the source file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue; // Skip missing files
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(filePath);

                // Copy each worksheet from the source workbook to the master workbook
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    // Add a copy of the worksheet to the master workbook using the sheet name
                    masterWorkbook.Worksheets.AddCopy(sheet.Name);
                }
            }

            // Define the output PDF file path
            string outputPdfPath = @"C:\Data\MergedWorkbook.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPdfPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the combined workbook as a single PDF preserving sheet order
            masterWorkbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine("Merged PDF created at: " + outputPdfPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
