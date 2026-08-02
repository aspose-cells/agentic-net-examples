// Title: Aspose.Cells for .NET – Create an Excel report of OLE object name, label, class ID and size
// Description: Loads a workbook, adds a sheet called "OLE Report", writes headers, then scans every worksheet (except the report) to extract each OLE object's Name, Label, ClassIdentifier (hex) and embedded data length. The collected data is written row‑by‑row and the workbook is saved with the new report.
// Keywords: Aspose.Cells | C# | .NET | OLE object report | extract OLE metadata | ClassIdentifier hex | embedded OLE size | Excel automation | list OleObject properties | generate worksheet report
// Common Searches: Aspose.Cells list OLE objects in workbook | C# extract OLE object name and size from Excel | how to get OLE class identifier using Aspose.Cells | create OLE metadata sheet with Aspose.Cells | enumerate embedded OLE objects in .xlsx
// Developer Intent: Produce an Excel sheet that enumerates every embedded OLE object's name, label, class identifier (hex) and file size.
// Use Cases: Compliance audit of embedded OLE content by exporting its metadata. | Estimating storage impact of OLE objects across a workbook. | Filtering OLE objects by class identifier for batch replacement or removal.
// AI Prompts: Add a column that shows the source file path for linked OLE objects. | Sort the generated OLE report by file size in descending order after populating the sheet. | Explain how to handle non‑embedded (linked) OLE objects when building the report with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectReportDemo
{
    // Loads a workbook, adds a sheet called "OLE Report", writes headers, then scans every worksheet (except the report) to extract each OLE object's Name, Label, ClassIdentifier (hex) and embedded data length. The collected data is written row‑by‑row and the workbook is saved with the new report.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the report
            int reportSheetIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
            reportSheet.Name = "OLE Report";

            // Write header titles
            reportSheet.Cells[0, 0].PutValue("Name");
            reportSheet.Cells[0, 1].PutValue("Label");
            reportSheet.Cells[0, 2].PutValue("Class Identifier");
            reportSheet.Cells[0, 3].PutValue("File Size (bytes)");

            int currentRow = 1; // start after header

            // Iterate through all worksheets and their OLE objects
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the report sheet itself to avoid self‑reference
                if (ws == reportSheet) continue;

                foreach (OleObject ole in ws.OleObjects)
                {
                    // Name of the OLE object
                    string name = ole.Name;

                    // Display label
                    string label = ole.Label;

                    // Class identifier as a hex string (handle null)
                    string classIdHex = ole.ClassIdentifier != null
                        ? BitConverter.ToString(ole.ClassIdentifier).Replace("-", "")
                        : string.Empty;

                    // File size – use embedded object data length if available
                    int fileSize = ole.ObjectData != null ? ole.ObjectData.Length : 0;

                    // Write the collected information into the report sheet
                    reportSheet.Cells[currentRow, 0].PutValue(name);
                    reportSheet.Cells[currentRow, 1].PutValue(label);
                    reportSheet.Cells[currentRow, 2].PutValue(classIdHex);
                    reportSheet.Cells[currentRow, 3].PutValue(fileSize);

                    currentRow++;
                }
            }

            // Save the workbook with the report (replace with desired output path)
            string outputPath = "WorkbookWithOleReport.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
