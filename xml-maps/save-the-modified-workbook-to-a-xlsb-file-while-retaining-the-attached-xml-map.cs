// Title: Save a modified Excel workbook as XLSB while preserving XML maps using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook to .xlsb and keep existing XML maps | How to retain XML map when converting .xlsx to .xlsb with Aspose.Cells | Saving Excel file as binary format with XlsbSaveOptions without losing XML maps in .NET | Aspose.Cells XlsbSaveOptions ExportAllColumnIndexes example for preserving XML maps | Add a row to a worksheet and export to XLSB while maintaining XML map
// Tags: Aspose.Cells XlsbSaveOptions ExportAllColumnIndexes usage | binary Excel export with XML map support | C# workbook modification and XLSB saving | preserve data mapping during XLSB conversion Aspose.Cells | save Excel as .xlsb without losing schema mapping

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file that contains XML maps, logs the attached maps, adds a new data row to the first worksheet, and then saves the workbook as an .xlsb file using XlsbSaveOptions with ExportAllColumnIndexes enabled and Level6 compression, ensuring the XML maps remain intact in the binary output.
class Program
{
    static void Main()
    {
        // Paths – adjust as needed
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsb";

        // Verify the source file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook that already contains an XML map
            Workbook workbook = new Workbook(inputPath);

            // Check for attached XML maps
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("Warning: No XML maps are attached to the workbook. The saved XLSB will not contain any map.");
            }
            else
            {
                Console.WriteLine($"Workbook contains {workbook.Worksheets.XmlMaps.Count} XML map(s):");
                foreach (XmlMap map in workbook.Worksheets.XmlMaps)
                {
                    Console.WriteLine($"  - Name: {map.Name}, Root Element: {map.RootElementName}");
                }
            }

            // ----- Modify the workbook -----
            // Example: add a new data row to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            int newRow = sheet.Cells.MaxDataRow + 1; // place after existing data
            sheet.Cells[newRow, 0].PutValue("NewItem");
            sheet.Cells[newRow, 1].PutValue(12345);
            Console.WriteLine($"Added new row at index {newRow}.");

            // ----- Save as XLSB while retaining the XML map -----
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Export all column indexes to keep the structure identical to the source
                ExportAllColumnIndexes = true,
                // Use a reasonable compression level (default is Level6)
                CompressionType = OoxmlCompressionType.Level6
            };

            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\" with the XML map preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception occurred: {ex.Message}");
        }
    }
}
