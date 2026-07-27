using System;
using Aspose.Cells;

class PdfExportWithRetry
{
    static void Main()
    {
        // Create a new workbook (use the provided create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["A2"].PutValue(123);
        worksheet.Cells["A3"].PutValue(DateTime.Now);

        // Configure PDF save options
        PdfSaveOptions options = new PdfSaveOptions
        {
            // Initial attempt: do not output a blank page when nothing to print
            OutputBlankPageWhenNothingToPrint = false
        };

        const string outputPath = "ExportedDocument.pdf";
        bool saved = false;

        try
        {
            // First attempt to save PDF (use the provided save rule)
            workbook.Save(outputPath, options);
            saved = true;
            Console.WriteLine("PDF saved successfully on first attempt.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"First PDF export failed: {ex.Message}");
        }

        if (!saved)
        {
            // Retry with OutputBlankPageWhenNothingToPrint set to true
            options.OutputBlankPageWhenNothingToPrint = true;
            try
            {
                workbook.Save(outputPath, options);
                Console.WriteLine("PDF saved successfully after retry with OutputBlankPageWhenNothingToPrint enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Retry PDF export also failed: {ex.Message}");
                // Further error handling could be placed here
            }
        }
    }
}

// Author: Aspose.Cells .NET example implementation.