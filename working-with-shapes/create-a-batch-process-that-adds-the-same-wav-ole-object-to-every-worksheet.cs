using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddWavOleToAllSheets
{
    static void Main()
    {
        // Paths for the workbook and the WAV file
        string workbookPath = "input.xlsx";
        string wavFilePath = "sample.wav";

        try
        {
            // Load existing workbook or create a new one if the file does not exist
            Workbook workbook;
            if (File.Exists(workbookPath))
                workbook = new Workbook(workbookPath);
            else
                workbook = new Workbook(); // creates a new empty workbook

            // Ensure the WAV file exists before reading it
            if (!File.Exists(wavFilePath))
                throw new FileNotFoundException($"WAV file not found: {wavFilePath}");

            // Read the WAV file into a byte array (object data)
            byte[] wavData = File.ReadAllBytes(wavFilePath);

            // Optional preview image for the OLE object (empty in this case)
            byte[] previewImage = new byte[0];

            // Add the same WAV OLE object to every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add an OLE object placeholder at row 5, column 5 with size 100x100 pixels
                int oleIndex = sheet.OleObjects.Add(5, 5, 100, 100, previewImage);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Embed the WAV file data (linkToFile = false)
                ole.SetEmbeddedObject(false, wavData, wavFilePath, false, "Audio");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}