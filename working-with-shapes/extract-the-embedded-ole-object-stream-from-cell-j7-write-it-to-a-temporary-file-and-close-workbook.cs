// Title: Extract an embedded OLE object from cell J7 in an Excel file using Aspose.Cells for .NET and save it as a temporary binary file
// AI Prompts: Write C# code with Aspose.Cells that finds the OLE object anchored at J7, reads its OleObjectData stream, and writes the bytes to a uniquely named temporary .bin file. | Show how to use reflection in .NET to access the OleObjectData property of an Aspose.Cells OleObject and export the embedded OLE content to disk.
// Common Searches: aspnet extract OLE object from specific cell in Excel using Aspose.Cells | c# save embedded OLE object from worksheet to temporary file | how to get OleObjectData stream from cell J7 with Aspose.Cells | retrieve binary data of an OLE object in Excel via Aspose.Cells .NET
// Tags: Aspose.Cells OLE extraction workflow | C# export binary data to temporary storage | Excel cell J7 OLE lookup | using reflection with Aspose.Cells objects | embedded OLE data export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads input.xlsx, locates the OLE object anchored at cell J7, uses reflection to obtain its OleObjectData stream, writes the binary data to a uniquely named temporary .bin file, and closes the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            using (var workbook = new Workbook(inputPath))
            {
                // Get the first worksheet (or specify the required one)
                var worksheet = workbook.Worksheets[0];

                // J7 => zero‑based row 6, column 9
                const int targetRow = 6;
                const int targetColumn = 9;

                // Locate the OLE object anchored at J7
                OleObject ole = null;
                foreach (OleObject obj in worksheet.OleObjects)
                {
                    if (obj.UpperLeftRow == targetRow && obj.UpperLeftColumn == targetColumn)
                    {
                        ole = obj;
                        break;
                    }
                }

                if (ole != null)
                {
                    try
                    {
                        // Use reflection to access OleObjectData (covers API variations)
                        var oleDataProp = ole.GetType().GetProperty("OleObjectData");
                        if (oleDataProp == null)
                        {
                            Console.WriteLine("OleObjectData property not found on OleObject.");
                            return;
                        }

                        var oleData = oleDataProp.GetValue(ole);
                        var getDataMethod = oleData?.GetType().GetMethod("GetData");
                        if (getDataMethod == null)
                        {
                            Console.WriteLine("GetData method not found on OleObjectData.");
                            return;
                        }

                        // Retrieve the OLE object's binary data
                        byte[] data = (byte[])getDataMethod.Invoke(oleData, null);

                        // Write the OLE stream to a temporary file
                        string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".bin");
                        File.WriteAllBytes(tempFile, data);
                        Console.WriteLine("OLE object extracted to: " + tempFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Failed to extract OLE data: " + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("No OLE object found in cell J7.");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
