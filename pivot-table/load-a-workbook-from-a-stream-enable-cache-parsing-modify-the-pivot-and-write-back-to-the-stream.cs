// Title: Load an XLSX workbook from a stream with cache parsing, rename the first pivot table row field, refresh it, and save the workbook back to a stream using Aspose.Cells for .NET
// Common Searches: c# aspnet load xlsx from memory stream with cache parsing aspose.cells | how to rename a pivot table row field using Aspose.Cells in C# | refresh and recalculate pivot table after changing field name with Aspose.Cells | save modified workbook to a MemoryStream using Aspose.Cells .NET | process workbook stream and output stream aspose.cells pivot table example
// Tags: load workbook from stream cache parsing | pivot table row field rename | pivot table refresh calculate Aspose.Cells | write workbook to memory stream | Aspose.Cells pivot table modification C#

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an XLSX workbook from an input Stream with cache‑parsing LoadOptions, locating the first pivot table, renaming its first row field, refreshing and recalculating the pivot data, and then saving the updated workbook to an output Stream in XLSX format. The code also resets the output stream position for downstream consumption.
public class PivotModifier
{
    /// <param name="inputStream">Stream containing the source workbook.</param>
    /// <param name="outputStream">Stream that will receive the modified workbook.</param>
    public void Process(Stream inputStream, Stream outputStream)
    {
        try
        {
            // Load the workbook from the input stream
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputStream, loadOptions);

            // Ensure there is at least one worksheet and one pivot table
            if (workbook.Worksheets.Count > 0 && workbook.Worksheets[0].PivotTables.Count > 0)
            {
                // Get the first pivot table in the first worksheet
                var pivotTable = workbook.Worksheets[0].PivotTables[0];

                // Example modification: rename the first row field (if it exists)
                if (pivotTable.RowFields.Count > 0)
                {
                    // Change the display name of the first row field
                    pivotTable.RowFields[0].Name = "ModifiedField";
                }

                // Refresh the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }

            // Save the modified workbook back to the output stream in XLSX format
            workbook.Save(outputStream, SaveFormat.Xlsx);
            // Reset the position of the output stream for downstream consumers
            outputStream.Position = 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }
}

public static class Program
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(inputPath, SaveFormat.Xlsx);
            }

            using (var inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var outputStream = new MemoryStream())
            {
                var modifier = new PivotModifier();
                modifier.Process(inputStream, outputStream);

                // Write the modified workbook to the output file
                using (var fileOut = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    outputStream.WriteTo(fileOut);
                }
            }

            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
