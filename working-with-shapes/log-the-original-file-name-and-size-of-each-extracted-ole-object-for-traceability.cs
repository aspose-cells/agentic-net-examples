using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ExtractOleObjects
{
    static void Main()
    {
        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            OleObjectCollection oleObjects = sheet.OleObjects;

            // Process each OLE object in the worksheet
            for (int i = 0; i < oleObjects.Count; i++)
            {
                OleObject ole = oleObjects[i];
                string sourceName;
                long sizeInBytes;

                if (ole.IsLink)
                {
                    // Linked OLE object: use the source file name
                    sourceName = ole.ObjectSourceFullName;

                    // Attempt to get the actual file size if the file exists on disk
                    sizeInBytes = 0;
                    if (File.Exists(sourceName))
                    {
                        sizeInBytes = new FileInfo(sourceName).Length;
                    }
                }
                else
                {
                    // Embedded OLE object: assign a placeholder name
                    sourceName = $"EmbeddedObject_{i}";

                    // Use ObjectData if available; otherwise fall back to FullObjectBin
                    byte[] data = ole.ObjectData ?? ole.FullObjectBin;
                    sizeInBytes = data?.Length ?? 0;
                }

                // Log the information
                Console.WriteLine($"Worksheet: {sheet.Name}, OLE Index: {i}, Source: {sourceName}, Size: {sizeInBytes} bytes");
            }
        }

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}