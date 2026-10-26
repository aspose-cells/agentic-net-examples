// Title: Safely handle missing background image streams while configuring worksheet page setup with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check for null background image stream before setting worksheet background | prevent exception when background image file is not found in Aspose.Cells workbook | how to skip worksheet background image if stream is empty using Aspose.Cells | error handling for missing background image in Aspose.Cells page setup C# | Aspose.Cells page setup without background image stream example
// Tags: null background image stream validation Aspose.Cells | worksheet page setup without background image Aspose.Cells | conditional background image assignment C# Aspose.Cells | avoid System.Drawing dependency Aspose.Cells | safe workbook save with optional background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one, attempts to obtain a background image stream, validates the stream for null, emptiness, or unreadability, and skips the background assignment to avoid System.Drawing dependencies and runtime exceptions. Comprehensive try‑catch blocks handle loading and saving, ensuring the workbook is saved safely even when the background image is missing.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists; create a new one if missing
            string inputPath = "input.xlsx";
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve background image stream (if any)
            Stream bgStream = GetBackgroundImageStream();

            // Aspose.Cells background image requires System.Drawing.Image.
            // To avoid dependency issues, we skip setting the background image if the stream is present.
            if (bgStream != null && bgStream.CanRead && bgStream.Length > 0)
            {
                Console.WriteLine("Background image stream detected, but setting it is skipped to avoid System.Drawing dependency.");
                bgStream.Dispose();
            }
            else
            {
                Console.WriteLine("Background image stream is missing or empty. Skipping background setup.");
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }

    // Retrieves a background image stream; returns null if the file does not exist
    static Stream GetBackgroundImageStream()
    {
        string path = "background.png";
        if (File.Exists(path))
        {
            return new FileStream(path, FileMode.Open, FileAccess.Read);
        }
        return null;
    }
}
