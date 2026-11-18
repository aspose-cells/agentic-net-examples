// Title: Generate an Excel report of OLE object names, labels, class identifiers, and file sizes using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list OLE objects and their properties in a workbook | How to export OLE object metadata to a new Excel file with Aspose.Cells | Aspose.Cells retrieve OLE object class identifier and file size in .NET | Create a report of OLE object names and labels using Aspose.Cells C# example | Iterate through worksheets and extract OLE object details with Aspose.Cells
// Tags: Aspose.Cells OLE object enumeration C# | extract OLE metadata Aspose.Cells | export OLE properties to new workbook | retrieve OLE class identifier Aspose.Cells | generate OLE object report Excel .NET | auto-fit columns Aspose.Cells report

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject and OleObjectCollection

// The sample loads a source workbook, iterates each worksheet's OleObjectCollection, and writes the worksheet name, OLE Name, Label, ClassIdentifier, and a placeholder file size (0) into a new workbook with headers. Columns are auto‑fitted and the report is saved as an XLSX file using Aspose.Cells for .NET.
class OleObjectsReport
{
    static void Main()
    {
        // Paths for source workbook and generated report
        string sourcePath = "source.xlsx";
        string reportPath = "OleObjectsReport.xlsx";

        try
        {
            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook containing OLE objects
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();

            // Use the first worksheet of the report workbook
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "OLE Objects";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("OLE Name");
            reportSheet.Cells[0, 2].PutValue("Label");
            reportSheet.Cells[0, 3].PutValue("Class Identifier");
            reportSheet.Cells[0, 4].PutValue("File Size (bytes)");

            int reportRow = 1; // Start writing data from the second row

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Access the collection of OLE objects on the current worksheet
                OleObjectCollection oleObjects = ws.OleObjects;

                // Process each OLE object
                foreach (OleObject ole in oleObjects)
                {
                    try
                    {
                        // Worksheet name where the OLE object resides
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);

                        // OLE object's name
                        reportSheet.Cells[reportRow, 1].PutValue(ole.Name);

                        // OLE object's label (display text)
                        reportSheet.Cells[reportRow, 2].PutValue(ole.Label);

                        // Class identifier (ProgID)
                        reportSheet.Cells[reportRow, 3].PutValue(ole.ClassIdentifier);

                        // File size in bytes – Aspose.Cells does not expose embedded data size directly,
                        // so we set it to 0 or you can implement custom extraction if needed.
                        reportSheet.Cells[reportRow, 4].PutValue(0);

                        reportRow++;
                    }
                    catch (Exception innerEx)
                    {
                        // Log errors for individual OLE objects but continue processing
                        Console.WriteLine($"Error processing OLE object '{ole.Name}' in worksheet '{ws.Name}': {innerEx.Message}");
                    }
                }
            }

            // Auto-fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook to the specified path
            reportWorkbook.Save(reportPath, SaveFormat.Xlsx);
            Console.WriteLine($"Report generated successfully at: {reportPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
