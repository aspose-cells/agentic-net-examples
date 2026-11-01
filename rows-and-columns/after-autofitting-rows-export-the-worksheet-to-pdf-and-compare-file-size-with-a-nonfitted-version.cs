// Title: Auto‑fit rows in an Excel worksheet with Aspose.Cells for .NET, export to PDF, and compare the resulting file sizes
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, calls AutoFitRows on the first worksheet, saves the workbook as a PDF, and prints the PDF file size in bytes. | Write a C# console application that loads the same workbook twice, exports one copy after applying AutoFitRows and another without any row adjustments, then outputs which PDF is smaller. | Create a C# example using Aspose.Cells to measure and compare the byte sizes of PDFs generated from a fitted worksheet versus a non‑fitted worksheet.
// Common Searches: Aspose.Cells .NET auto fit rows before PDF export and check file size | compare PDF size of Excel sheet with and without AutoFitRows using C# | measure impact of row auto‑fit on PDF output size in Aspose.Cells
// Tags: AutoFitRows PDF export Aspose.Cells | row auto‑fit size impact PDF Aspose.Cells | compare fitted vs non‑fitted PDF Aspose.Cells | measure PDF file size Aspose.Cells .NET | Excel to PDF conversion row height optimization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, auto‑fits all rows in the first worksheet, saves it as a PDF, then reloads the original workbook without auto‑fit, saves a second PDF, and prints both file sizes along with a message indicating which PDF is smaller.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourceFile = "input.xlsx";

        // Load the workbook (using the provided load rule)
        Workbook wbOriginal = new Workbook(sourceFile);

        // ------------------- Auto‑fit rows and export -------------------
        // Get the first worksheet
        Worksheet wsFitted = wbOriginal.Worksheets[0];

        // Auto‑fit all rows in the worksheet
        wsFitted.AutoFitRows();

        // Export the auto‑fitted worksheet to PDF (using the provided save rule)
        string fittedPdfPath = "fitted.pdf";
        wbOriginal.Save(fittedPdfPath, SaveFormat.Pdf);

        // Get the file size of the fitted PDF
        long fittedSize = new FileInfo(fittedPdfPath).Length;

        // ------------------- Export without auto‑fit -------------------
        // Reload the original workbook to get a clean copy (no auto‑fit applied)
        Workbook wbNonFitted = new Workbook(sourceFile);

        // Export the non‑fitted worksheet to PDF
        string nonFittedPdfPath = "nonfitted.pdf";
        wbNonFitted.Save(nonFittedPdfPath, SaveFormat.Pdf);

        // Get the file size of the non‑fitted PDF
        long nonFittedSize = new FileInfo(nonFittedPdfPath).Length;

        // ------------------- Compare file sizes -------------------
        Console.WriteLine($"Fitted PDF size: {fittedSize} bytes");
        Console.WriteLine($"Non‑fitted PDF size: {nonFittedSize} bytes");

        if (fittedSize < nonFittedSize)
        {
            Console.WriteLine("Auto‑fitting rows reduced the PDF file size.");
        }
        else if (fittedSize > nonFittedSize)
        {
            Console.WriteLine("Auto‑fitting rows increased the PDF file size.");
        }
        else
        {
            Console.WriteLine("PDF file sizes are identical.");
        }
    }
}
