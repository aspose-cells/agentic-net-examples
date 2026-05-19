using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectReport
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains OLE objects
            // (Replace the path with the actual file location)
            Workbook workbook = new Workbook("input.xlsx");

            // Create a new worksheet to hold the report
            int reportSheetIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
            reportSheet.Name = "OLE Report";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("OLE Name");
            reportSheet.Cells[0, 2].PutValue("Label");
            reportSheet.Cells[0, 3].PutValue("Class Identifier (Hex)");
            reportSheet.Cells[0, 4].PutValue("File Size (bytes)");

            int currentRow = 1;

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Access the collection of OLE objects in the current worksheet
                OleObjectCollection oleObjects = ws.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Retrieve required properties
                    string oleName = ole.Name;
                    string label = ole.Label;
                    byte[] classIdBytes = ole.ClassIdentifier;
                    string classIdHex = classIdBytes != null ? BitConverter.ToString(classIdBytes).Replace("-", "") : string.Empty;

                    // Determine file size: use embedded data if present, otherwise 0
                    long fileSize = ole.ObjectData != null ? ole.ObjectData.Length : 0;

                    // Write data to the report sheet
                    reportSheet.Cells[currentRow, 0].PutValue(ws.Name);
                    reportSheet.Cells[currentRow, 1].PutValue(oleName);
                    reportSheet.Cells[currentRow, 2].PutValue(label);
                    reportSheet.Cells[currentRow, 3].PutValue(classIdHex);
                    reportSheet.Cells[currentRow, 4].PutValue(fileSize);

                    currentRow++;
                }
            }

            // Save the workbook with the added report worksheet
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}