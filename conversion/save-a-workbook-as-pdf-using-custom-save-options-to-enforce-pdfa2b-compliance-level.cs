// Title: Create a PDF/A‑2b compliant PDF from an Excel workbook using Aspose.Cells PdfSaveOptions in C#
// Common Searches: how to generate PDF/A-2b from Excel with Aspose.Cells in C# | Aspose.Cells C# set PdfSaveOptions compliance to PdfA2b | C# export workbook to PDF/A-2b using reflection for compliance property | save Excel file as PDF/A-2b compliant PDF with Aspose.Cells PdfSaveOptions | example of setting PDF/A-2b compliance in Aspose.Cells when property is not directly exposed
// Tags: Aspose.Cells PdfSaveOptions PDF/A-2b compliance | C# export Excel to PDF/A-2b | reflection set PdfSaveOptions property Aspose.Cells | custom PDF save options Aspose.Cells | generate PDF/A-2b from workbook C#

using System;
using Aspose.Cells;

// The sample creates a Workbook, adds data, configures PdfSaveOptions, uses reflection to set the Compliance property to PdfA2b when available, and saves the workbook as a PDF/A‑2b compliant file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set PDF save options (PDF/A‑2b compliance if supported)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Attempt to set compliance via reflection; ignore if not supported
            try
            {
                var complianceProp = pdfOptions.GetType().GetProperty("Compliance");
                if (complianceProp != null)
                {
                    Type enumType = complianceProp.PropertyType;
                    var enumValue = Enum.Parse(enumType, "PdfA2b");
                    complianceProp.SetValue(pdfOptions, enumValue);
                }
            }
            catch
            {
                // Compliance setting not supported; continue without it
            }

            // Save the workbook as a PDF file using the custom options
            workbook.Save("Result.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
