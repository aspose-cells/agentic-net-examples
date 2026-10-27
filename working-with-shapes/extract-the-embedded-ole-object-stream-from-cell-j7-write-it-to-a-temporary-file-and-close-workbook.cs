// Title: Extract an embedded OLE object from cell J7 and save it to a temporary file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells extract OLE object from a specific worksheet cell | How to retrieve raw OLE data from Excel using Aspose.Cells | Save embedded OLE object to a temporary file in .NET | Get OleObject stream from cell J7 with Aspose.Cells | Extract binary OLE content from an Excel workbook using C#
// Tags: OleObject raw data extraction using Aspose.Cells | write embedded OLE stream to temporary file in C# | retrieve binary OLE content from Excel worksheet | Aspose.Cells handling of OLE objects | temporary file creation for extracted OLE data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

// The example loads an Excel workbook with Aspose.Cells, locates the OleObject anchored at cell J7, reads its ObjectData byte array, writes the binary stream to a uniquely named temporary file, and then disposes the workbook.
class ExtractOleObject
{
    static void Main()
    {
        try
        {
            // Path to the workbook (adjust as needed)
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            using (Workbook workbook = new Workbook(workbookPath))
            {
                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Target cell J7 (zero‑based indices)
                int targetRow = 6;    // Row index for J7
                int targetColumn = 9; // Column index for J7

                // Find the OleObject anchored at J7
                OleObject ole = null;
                foreach (OleObject obj in sheet.OleObjects)
                {
                    if (obj.UpperLeftRow == targetRow && obj.UpperLeftColumn == targetColumn)
                    {
                        ole = obj;
                        break;
                    }
                }

                if (ole == null)
                {
                    Console.WriteLine("No OLE object found in cell J7.");
                }
                else
                {
                    try
                    {
                        // Extract raw OLE data
                        byte[] oleData = ole.ObjectData;

                        // Write data to a temporary file
                        string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".bin");
                        File.WriteAllBytes(tempFilePath, oleData);

                        Console.WriteLine($"OLE object extracted to: {tempFilePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to extract OLE object: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
