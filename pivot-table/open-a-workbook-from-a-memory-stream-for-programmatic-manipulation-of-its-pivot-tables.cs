// Title: Open an Excel workbook from a MemoryStream and enumerate its pivot tables using Aspose.Cells for .NET (C#)
// Common Searches: c# load excel file from byte array using Aspose.Cells | asp.net read excel from memory stream and retrieve pivot table names | aspose.cells get list of pivot tables after opening workbook from stream | programmatically enumerate pivot tables in a worksheet with Aspose.Cells | save updated excel workbook to a memory stream in C#
// Tags: load workbook from memory stream Aspose.Cells | access worksheet pivot tables Aspose.Cells | list pivot table names C# Aspose.Cells | save workbook to memory stream Aspose.Cells | modify pivot table fields Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotTableCollection

// The example reads an Excel file into a byte array, loads it into an Aspose.Cells Workbook via a MemoryStream, accesses the first worksheet, iterates through its PivotTableCollection to print each pivot table's name, and optionally saves the modified workbook back into a new MemoryStream.
class PivotTableExample
{
    static void Main()
    {
        const string inputFile = "SamplePivot.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Read the Excel file into a byte array (e.g., from a database or network)
            byte[] excelBytes = File.ReadAllBytes(inputFile);

            // Load the workbook from a memory stream
            using (MemoryStream memoryStream = new MemoryStream(excelBytes))
            {
                Workbook workbook = new Workbook(memoryStream);

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of pivot tables on the worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;

                // Iterate through all pivot tables and print their names
                foreach (PivotTable pivotTable in pivotTables)
                {
                    Console.WriteLine($"Pivot Table Name: {pivotTable.Name}");
                    // Additional manipulation can be performed here, e.g., refresh data, modify fields, etc.
                    // pivotTable.RefreshData();
                }

                // (Optional) Save the modified workbook back to a memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    // The outputStream now contains the updated Excel file.
                    // It can be written back to a file, sent over a network, etc.
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
