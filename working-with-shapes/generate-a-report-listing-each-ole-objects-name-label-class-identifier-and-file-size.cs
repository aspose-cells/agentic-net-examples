// Title: Create an Excel report of embedded OLE object details (name, label, class ID, file size) using Aspose.Cells for .NET
// Common Searches: how to enumerate embedded OLE objects in an Excel workbook with Aspose.Cells C# | Aspose.Cells C# extract OLE object name, label and class identifier | generate a spreadsheet listing OLE object file sizes using Aspose.Cells .NET | C# code to write OLE object metadata to a new Excel file with Aspose.Cells | retrieve embedded OLE data length from a worksheet using Aspose.Cells API
// Tags: enumerate OleObject collection Aspose.Cells | extract OLE metadata to Excel workbook | write OLE object details with Aspose.Cells .NET | calculate embedded OLE file size C# | generate OLE objects report Excel Aspose

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads a source workbook, iterates through each worksheet's OleObject collection, captures each object's Name, Label, ClassIdentifier, and the size of its embedded data (when accessible), writes these values into a new worksheet named "OLE Objects Report", and saves the report as a separate Excel file.
class OLEReportGenerator
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string reportPath = "OLEObjectsReport.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook containing OLE objects
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook for the OLE objects report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "OLE Objects Report";

            // Write column headers
            reportSheet.Cells[0, 0].PutValue("Name");
            reportSheet.Cells[0, 1].PutValue("Label");
            reportSheet.Cells[0, 2].PutValue("Class Identifier");
            reportSheet.Cells[0, 3].PutValue("File Size (bytes)");

            int currentRow = 1; // Start writing data from the second row

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Iterate through each OLE object in the current worksheet
                foreach (OleObject ole in ws.OleObjects)
                {
                    // Write OLE object details to the report sheet
                    reportSheet.Cells[currentRow, 0].PutValue(ole.Name);
                    reportSheet.Cells[currentRow, 1].PutValue(ole.Label);
                    reportSheet.Cells[currentRow, 2].PutValue(ole.ClassIdentifier);

                    // Determine the file size of the embedded OLE object (using dynamic to avoid compile‑time binding)
                    int fileSize = 0;
                    try
                    {
                        var oleData = ((dynamic)ole).OleObjectData;
                        if (oleData != null && oleData.Data != null)
                        {
                            fileSize = oleData.Data.Length;
                        }
                    }
                    catch
                    {
                        // Property not available in this version; leave fileSize as 0
                    }

                    reportSheet.Cells[currentRow, 3].PutValue(fileSize);
                    currentRow++;
                }
            }

            // Save the generated report to a new Excel file
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
