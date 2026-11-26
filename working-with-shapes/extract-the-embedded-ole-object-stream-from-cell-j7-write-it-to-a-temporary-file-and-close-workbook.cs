// Title: Extract an embedded OLE object from cell J7 and save it to a temporary file using Aspose.Cells for .NET
// Common Searches: C# code to extract an OLE object anchored at J7 from an Excel workbook with Aspose.Cells | How to write embedded OLE object data to a temporary file in Aspose.Cells .NET | Retrieve the binary stream of an OLE object located in a specific worksheet cell using Aspose.Cells for C#
// Tags: Aspose.Cells extract embedded OLE object C# | write OLE object stream to temporary file .NET | locate OLE object by cell address J7 Aspose.Cells | binary OLE data extraction C# | using block workbook disposal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, scans the first worksheet for an OLE object whose upper‑left corner is anchored at cell J7, obtains its ObjectData byte array, writes the bytes to a uniquely named temporary .bin file, and relies on a using block to automatically close and dispose the workbook.
class Program
{
    static void Main(string[] args)
    {
        // Validate input arguments
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program.exe <path-to-workbook>");
            return;
        }

        string workbookPath = args[0];

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" does not exist.");
            return;
        }

        // The target cell is J7 (zero‑based indices: column 9, row 6)
        const int targetRow = 6;      // Row 7
        const int targetColumn = 9;   // Column J

        try
        {
            // Load the workbook inside a using block to ensure proper disposal
            using (Workbook workbook = new Workbook(workbookPath))
            {
                // Assume the OLE object is on the first worksheet; adjust if needed
                Worksheet sheet = workbook.Worksheets[0];

                OleObject targetOle = null;

                // Iterate through all OLE objects on the worksheet to find the one anchored at J7
                foreach (OleObject ole in sheet.OleObjects)
                {
                    if (ole.UpperLeftRow == targetRow && ole.UpperLeftColumn == targetColumn)
                    {
                        targetOle = ole;
                        break;
                    }
                }

                if (targetOle == null)
                {
                    Console.WriteLine("No embedded OLE object found at cell J7.");
                }
                else
                {
                    // Retrieve the embedded object's binary data
                    byte[] oleData = targetOle.ObjectData;

                    if (oleData == null || oleData.Length == 0)
                    {
                        Console.WriteLine("The OLE object at J7 does not contain any data.");
                    }
                    else
                    {
                        // Create a temporary file name with appropriate extension if known
                        string tempFilePath = Path.Combine(Path.GetTempPath(),
                            $"OleObject_J7_{Guid.NewGuid():N}.bin");

                        File.WriteAllBytes(tempFilePath, oleData);
                        Console.WriteLine($"OLE object extracted successfully to: {tempFilePath}");
                    }
                }

                // Workbook will be closed automatically when exiting the using block
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
