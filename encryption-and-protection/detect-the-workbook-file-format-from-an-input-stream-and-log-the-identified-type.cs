// Title: Detect the Excel workbook format from a Stream using Aspose.Cells LoadOptions in C#
// Common Searches: Aspose.Cells C# detect workbook file type from memory stream | How to get FileFormatType after loading an Excel file with Aspose.Cells | Load Excel workbook from a stream and identify .xlsx or .xls using Aspose.Cells | C# code to auto‑detect Excel format when opening a workbook from a file stream | Retrieve detected workbook format with Aspose.Cells LoadOptions in .NET
// Tags: Aspose.Cells stream workbook format detection | LoadOptions auto detect Excel file type .NET | FileFormatType retrieval after loading workbook | C# detect .xlsx vs .xls using Aspose.Cells | exception handling for missing workbook stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example opens an Excel file as a Stream, uses default LoadOptions to enable automatic format detection, loads the Workbook with Aspose.Cells, reads the detected FileFormatType via workbook.FileFormat, and writes the result to the console while handling missing file and other exceptions.
class WorkbookFormatDetector
{
    static void Main()
    {
        try
        {
            // Obtain the workbook stream (file, memory, etc.).
            using Stream inputStream = GetWorkbookStream();

            // Use default LoadOptions to enable auto‑detection of the format.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook from the stream with the specified options.
            Workbook workbook = new Workbook(inputStream, loadOptions);

            // Retrieve the detected file format.
            FileFormatType detectedFormat = workbook.FileFormat;

            // Output the identified format.
            Console.WriteLine($"Detected workbook format: {detectedFormat}");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File not found: {ex.FileName}");
        }
        catch (NotImplementedException ex)
        {
            Console.Error.WriteLine($"Implementation missing: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieves a stream containing workbook data.
    // Attempts to open "sample.xlsx" from the application directory.
    private static Stream GetWorkbookStream()
    {
        const string fileName = "sample.xlsx";

        if (!File.Exists(fileName))
        {
            // If the sample file is missing, inform the caller.
            throw new FileNotFoundException("The required workbook file was not found.", fileName);
        }

        // Open the file for reading and return the stream.
        return new FileStream(fileName, FileMode.Open, FileAccess.Read);
    }
}
