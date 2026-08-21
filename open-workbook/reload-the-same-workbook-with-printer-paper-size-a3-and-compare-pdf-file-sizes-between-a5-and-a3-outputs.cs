// Title: Reload Workbook with A3 & A5 Paper Sizes and Compare PDF File Sizes – Aspose.Cells for .NET
// Description: C# example that creates a sample workbook, saves it as XLSX, then reloads the file twice using LoadOptions.SetPaperSize – once with PaperA5 and once with PaperA3. Each version is exported to PDF, the byte lengths are read, and a concise comparison is printed, demonstrating how paper size affects PDF size in Aspose.Cells.
// Keywords: Aspose.Cells | LoadOptions.SetPaperSize | PaperSizeType.PaperA5 | PaperSizeType.PaperA3 | PDF conversion .NET | compare PDF file size | C# workbook reload | Excel to PDF Aspose
// Common Searches: Aspose.Cells set paper size for PDF conversion | C# compare PDF size A5 vs A3 | LoadOptions paper size example Aspose.Cells | How to reload Excel workbook with different paper size | Measure PDF output size Aspose.Cells .NET
// Developer Intent: Reload the same Excel workbook with A5 and A3 paper settings, export each to PDF, and determine which paper size produces the larger file.
// Use Cases: Validate storage impact of different paper formats before bulk PDF generation. | Automate quality checks that ensure PDF size stays within limits for specific print layouts. | Create side‑by‑side PDFs in A5 and A3 for reporting, then log their sizes for performance analysis.
// AI Prompts: Generate C# code that loads an existing Excel file with LoadOptions.SetPaperSize(PaperSizeType.PaperA4) and saves it as a PDF using Aspose.Cells. | Explain why changing the paper size during workbook loading influences the resulting PDF size and suggest compression settings to minimize file size. | Write a C# routine that iterates over a list of PaperSizeType values, creates PDFs for each, and outputs a table of file sizes.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a sample workbook, saves it as XLSX, then reloads the file twice using LoadOptions.SetPaperSize – once with PaperA5 and once with PaperA3. Each version is exported to PDF, the byte lengths are read, and a concise comparison is printed, demonstrating how paper size affects PDF size in Aspose.Cells.
class ComparePdfFileSizes
{
    static void Main()
    {
        // Paths for temporary and output files
        string xlsxPath = "sample.xlsx";
        string pdfA5Path = "sample_A5.pdf";
        string pdfA3Path = "sample_A3.pdf";

        // -------------------------------------------------
        // 1. Create a workbook with some sample data
        // -------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate a simple table
        ws.Cells["A1"].PutValue("Item");
        ws.Cells["B1"].PutValue("Quantity");
        for (int i = 2; i <= 51; i++)
        {
            ws.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            ws.Cells[$"B{i}"].PutValue((i - 1) * 5);
        }

        // Save the workbook to a temporary .xlsx file (used for re‑loading)
        wb.Save(xlsxPath, SaveFormat.Xlsx);

        // -------------------------------------------------
        // 2. Load workbook with A5 paper size and save as PDF
        // -------------------------------------------------
        LoadOptions loadA5 = new LoadOptions();
        loadA5.SetPaperSize(PaperSizeType.PaperA5);               // rule: LoadOptions.SetPaperSize
        Workbook wbA5 = new Workbook(xlsxPath, loadA5);
        wbA5.Save(pdfA5Path, SaveFormat.Pdf);                    // rule: Workbook.Save

        long sizeA5 = new FileInfo(pdfA5Path).Length;

        // -------------------------------------------------
        // 3. Load the same workbook with A3 paper size and save as PDF
        // -------------------------------------------------
        LoadOptions loadA3 = new LoadOptions();
        loadA3.SetPaperSize(PaperSizeType.PaperA3);               // rule: LoadOptions.SetPaperSize
        Workbook wbA3 = new Workbook(xlsxPath, loadA3);
        wbA3.Save(pdfA3Path, SaveFormat.Pdf);                    // rule: Workbook.Save

        long sizeA3 = new FileInfo(pdfA3Path).Length;

        // -------------------------------------------------
        // 4. Compare file sizes and output the result
        // -------------------------------------------------
        Console.WriteLine($"PDF size with A5 paper: {sizeA5} bytes");
        Console.WriteLine($"PDF size with A3 paper: {sizeA3} bytes");

        if (sizeA3 > sizeA5)
            Console.WriteLine("A3 PDF is larger than A5 PDF.");
        else if (sizeA3 < sizeA5)
            Console.WriteLine("A3 PDF is smaller than A5 PDF.");
        else
            Console.WriteLine("Both PDFs have the same size.");
    }
}
