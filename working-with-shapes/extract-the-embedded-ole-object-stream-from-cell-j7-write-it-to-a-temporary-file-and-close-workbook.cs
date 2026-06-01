using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractOleObject
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Access the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the OLE object that is positioned at cell J7 (row index 6, column index 9)
        OleObject targetOle = null;
        foreach (OleObject ole in sheet.OleObjects)
        {
            if (ole.UpperLeftRow == 6 && ole.UpperLeftColumn == 9)
            {
                targetOle = ole;
                break;
            }
        }

        if (targetOle == null)
        {
            Console.WriteLine("No OLE object found at cell J7.");
        }
        else
        {
            // Retrieve the embedded OLE data (use ObjectData or FullObjectBin as needed)
            byte[] oleData = targetOle.ObjectData ?? targetOle.FullObjectBin;

            if (oleData != null && oleData.Length > 0)
            {
                // Create a temporary file and write the OLE data to it
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".bin");
                File.WriteAllBytes(tempFilePath, oleData);
                Console.WriteLine($"OLE object data extracted to: {tempFilePath}");
            }
            else
            {
                Console.WriteLine("OLE object contains no data.");
            }
        }

        // Close the workbook (dispose resources)
        workbook.Dispose();
    }
}