// Title: Refresh all slicers and pivot tables in an Excel workbook loaded from a MemoryStream and export to PDF using Aspose.Cells for .NET
// AI Prompts: Create a C# method that takes a MemoryStream with an .xlsx file, refreshes every slicer and all pivot tables in the workbook, and returns a PDF as a MemoryStream using Aspose.Cells. | Extend the conversion routine to accept an optional list of slicer names, refresh only those slicers, then save the workbook to a PDF stream with Aspose.Cells.
// Common Searches: how to refresh slicers in an Excel file before converting to PDF with Aspose.Cells C# | convert Excel workbook from MemoryStream to PDF while updating pivot tables using Aspose.Cells .NET | Aspose.Cells programmatically refresh all slicers C# | save refreshed Excel workbook as PDF without writing an intermediate file Aspose.Cells
// Tags: Aspose.Cells update slicer selections | Aspose.Cells programmatic pivot table refresh | in-memory Excel to PDF conversion Aspose.Cells | slicer handling before PDF export Aspose.Cells | memory stream workbook processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel workbook from a MemoryStream, refreshes every slicer and all pivot tables, then saves the workbook as a PDF into a new MemoryStream using Aspose.Cells.
public static class WorkbookProcessor
{
    /// <param name="excelStream">MemoryStream containing the source Excel file.</param>
    /// <returns>MemoryStream containing the PDF output.</returns>
    public static MemoryStream ConvertToPdfWithRefresh(MemoryStream excelStream)
    {
        try
        {
            // Ensure the input stream is positioned at the beginning
            excelStream.Position = 0;

            // Load the workbook from the stream
            Workbook workbook = new Workbook(excelStream);

            // Refresh slicers (if any) to reflect current data selections.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 0; i < sheet.Slicers.Count; i++)
                {
                    Slicer slicer = sheet.Slicers[i];
                    slicer.Refresh();
                }
            }

            // Refresh all pivot tables in the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Save the refreshed workbook as PDF into a new memory stream
            MemoryStream pdfStream = new MemoryStream();
            workbook.Save(pdfStream, SaveFormat.Pdf);
            pdfStream.Position = 0;
            return pdfStream;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Read the Excel file into a memory stream
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            using (MemoryStream excelStream = new MemoryStream())
            {
                fileStream.CopyTo(excelStream);

                // Convert to PDF with slicer and pivot refresh
                MemoryStream pdfStream = WorkbookProcessor.ConvertToPdfWithRefresh(excelStream);

                // Write the PDF to disk
                using (FileStream outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    pdfStream.CopyTo(outFile);
                }

                Console.WriteLine($"PDF successfully created at: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
