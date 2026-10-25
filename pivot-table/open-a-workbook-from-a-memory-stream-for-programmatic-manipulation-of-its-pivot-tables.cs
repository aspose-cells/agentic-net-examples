// Title: Load an Excel workbook from a MemoryStream and rename/refresh its pivot tables using Aspose.Cells for .NET
// Common Searches: how to open an Excel file from a byte array with Aspose.Cells in C# | rename pivot tables programmatically with Aspose.Cells .NET | refresh pivot table data after modifying a workbook in memory using Aspose.Cells | save a workbook to a file after processing pivot tables in C#
// Tags: open workbook from memory stream Aspose.Cells | pivot table rename operation Aspose.Cells | pivot table refresh operation Aspose.Cells | save workbook to byte array C# | in‑memory pivot table manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable classes

// The example reads an existing Excel file (or creates a blank one if missing), loads it into a byte array, opens the workbook from a MemoryStream, accesses the first worksheet, iterates through all pivot tables to append "_Modified" to each name, refreshes and recalculates the pivot data, and finally saves the updated workbook to a new file via another MemoryStream.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a blank workbook as a fallback
            if (!File.Exists(inputPath))
            {
                // Create a new workbook with a single worksheet
                var newWb = new Workbook();
                newWb.Worksheets[0].Name = "Sheet1";
                newWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the Excel file into a byte array
            byte[] excelBytes = File.ReadAllBytes(inputPath);

            // Open the workbook from a memory stream
            using (MemoryStream inputStream = new MemoryStream(excelBytes))
            {
                // Create a Workbook instance from the stream
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the collection of pivot tables on the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Example manipulation: rename each pivot table and refresh it
                foreach (PivotTable pt in pivotTables)
                {
                    // Append a suffix to the pivot table name
                    pt.Name = pt.Name + "_Modified";

                    // Refresh data and recalculate the pivot table after changes
                    pt.RefreshData();
                    pt.CalculateData();
                }

                // Save the modified workbook back to a memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    File.WriteAllBytes(outputPath, outputStream.ToArray());
                }
            }

            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
