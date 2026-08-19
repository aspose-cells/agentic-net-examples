// Title: Refresh Slicers & Pivot Tables and Export Excel to PDF from MemoryStream using Aspose.Cells (C#)
// Description: Loads an Excel workbook from a MemoryStream, refreshes every slicer (automatically updating linked pivot tables), forces a full pivot‑table refresh, and saves the result as a PDF into a new MemoryStream—all without touching the file system. Ideal for in‑memory reporting with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | MemoryStream | Excel to PDF | slicer refresh | pivot table refresh | in‑memory conversion | PDF export | Aspose.Cells Slicer | Aspose.Cells PivotTable
// Common Searches: Aspose.Cells refresh slicer before PDF export | C# convert Excel stream to PDF with updated pivot tables | How to refresh slicer selections programmatically in Aspose.Cells | Export Excel to PDF from MemoryStream without saving to disk | Refresh all slicers in a workbook using Aspose.Cells
// Developer Intent: Refresh all slicers and pivot tables in a workbook loaded from a MemoryStream and return the updated PDF as a MemoryStream.
// Use Cases: Web API that receives an Excel template as a byte array, applies slicer filters, and streams a PDF back to the client. | Scheduled service that programmatically sets slicer values, refreshes pivot tables, and generates PDF summaries for email distribution. | Desktop utility that converts user‑selected Excel files to PDF on‑the‑fly, keeping all processing in memory to avoid temporary files.
// AI Prompts: Generate C# code with Aspose.Cells to load an Excel file from a MemoryStream, refresh all slicers, refresh pivot tables, and save the workbook as a PDF to another MemoryStream. | Explain why Slicer.Refresh also updates linked pivot tables and what additional steps ensure the PDF reflects the latest slicer state. | Provide best‑practice error handling for converting an Excel workbook with slicers to PDF using streams in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel workbook from a MemoryStream, refreshes every slicer (automatically updating linked pivot tables), forces a full pivot‑table refresh, and saves the result as a PDF into a new MemoryStream—all without touching the file system. Ideal for in‑memory reporting with Aspose.Cells for .NET.
public static class WorkbookProcessor
{
    /// <param name="excelStream">MemoryStream containing the source Excel file.</param>
    /// <returns>MemoryStream containing the PDF output.</returns>
    public static MemoryStream ConvertToPdfWithSlicerRefresh(MemoryStream excelStream)
    {
        if (excelStream == null)
            throw new ArgumentNullException(nameof(excelStream));

        try
        {
            // Ensure the stream is positioned at the beginning
            if (excelStream.CanSeek)
                excelStream.Position = 0;

            // Load the workbook from the provided stream
            Workbook workbook = new Workbook(excelStream);

            // Refresh all slicers (if any) to update linked pivot tables
            foreach (Worksheet ws in workbook.Worksheets)
            {
                for (int i = 0; i < ws.Slicers.Count; i++)
                {
                    Slicer slicer = ws.Slicers[i];
                    slicer.Refresh(); // Refreshes the slicer and associated pivot tables
                }
            }

            // Refresh pivot tables that are not directly linked to slicers
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook as PDF into a new memory stream
            MemoryStream pdfStream = new MemoryStream();
            workbook.Save(pdfStream, SaveFormat.Pdf);
            pdfStream.Position = 0;

            return pdfStream;
        }
        catch (Exception ex)
        {
            // Wrap any exception for caller context
            throw new InvalidOperationException("Error converting Excel to PDF with slicer refresh.", ex);
        }
    }
}

// Dummy entry point to satisfy the compiler
public class Program
{
    public static void Main()
    {
        // Placeholder main method
    }
}
