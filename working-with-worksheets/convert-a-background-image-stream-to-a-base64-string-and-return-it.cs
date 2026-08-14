// Title: C# Aspose.Cells – Convert Worksheet Background Image to a Base64 String
// Description: Sample C# code that loads an Excel workbook with Aspose.Cells, extracts the background picture of a specified worksheet, and returns it as a Base64‑encoded string. Includes a reusable helper for converting any image stream to Base64 and robust checks for file existence and missing images.
// Keywords: Aspose.Cells background image | C# convert image to Base64 | Excel worksheet picture extraction | Base64 string from Excel image | .NET Excel image stream | retrieve worksheet background bytes | data URI Excel background | Aspose.Cells API example
// Common Searches: Aspose.Cells get worksheet background image as Base64 | C# extract Excel sheet background picture | convert Excel background to Base64 string | how to read worksheet background image with Aspose.Cells | Base64 encoding of Excel worksheet image
// Developer Intent: Obtain the background picture of an Excel worksheet and deliver it as a Base64‑encoded string for embedding or transmission.
// Use Cases: Embed the worksheet background directly in HTML or email using a data‑URI. | Store the Base64 image in a database for audit trails or later rendering. | Transmit the image via a REST API to a web client without saving a physical file.
// AI Prompts: Generate a method that opens an Excel file with Aspose.Cells, reads the background image of a given worksheet, and returns a Base64 string, handling missing files and empty backgrounds. | Extend the ConvertStreamToBase64 utility to accept non‑seekable streams and optionally prepend the appropriate MIME type for data‑URI usage. | Show how to take the Base64 string from GetWorksheetBackgroundImageBase64 and create an <img> tag with a data URI in an ASP.NET MVC view.

using System;
using System.IO;
using Aspose.Cells;

// Sample C# code that loads an Excel workbook with Aspose.Cells, extracts the background picture of a specified worksheet, and returns it as a Base64‑encoded string. Includes a reusable helper for converting any image stream to Base64 and robust checks for file existence and missing images.
public static class BackgroundImageHelper
{
    // Converts any image stream (e.g., a worksheet background image stream) to a Base64 string.
    public static string ConvertStreamToBase64(Stream imageStream)
    {
        if (imageStream == null)
            throw new ArgumentNullException(nameof(imageStream));

        // Reset position if the stream supports seeking.
        if (imageStream.CanSeek)
            imageStream.Position = 0;

        // Read the entire stream into a byte array.
        using (MemoryStream ms = new MemoryStream())
        {
            imageStream.CopyTo(ms);
            byte[] imageBytes = ms.ToArray();

            // Return the Base64 representation.
            return Convert.ToBase64String(imageBytes);
        }
    }

    // Loads a workbook, retrieves the background image of a worksheet,
    // and returns it as a Base64 string.
    public static string GetWorksheetBackgroundImageBase64(string workbookPath, int worksheetIndex = 0)
    {
        // Ensure the workbook file exists to avoid FileNotFoundException.
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the workbook.
        Workbook workbook = new Workbook(workbookPath);
        Worksheet worksheet = workbook.Worksheets[worksheetIndex];

        // Retrieve the background image bytes.
        byte[] backgroundBytes = worksheet.BackgroundImage;

        // If no background image is set, return an empty string.
        if (backgroundBytes == null || backgroundBytes.Length == 0)
            return string.Empty;

        // Convert the byte array to Base64.
        return Convert.ToBase64String(backgroundBytes);
    }
}

public class Program
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Example workbook path; adjust as needed.
        string workbookPath = "sample.xlsx";

        // Verify the file exists before proceeding.
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        try
        {
            // Retrieve the background image as a Base64 string.
            string base64Image = BackgroundImageHelper.GetWorksheetBackgroundImageBase64(workbookPath);

            if (string.IsNullOrEmpty(base64Image))
                Console.WriteLine("No background image found in the worksheet.");
            else
                Console.WriteLine($"Background Image Base64: {base64Image}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
