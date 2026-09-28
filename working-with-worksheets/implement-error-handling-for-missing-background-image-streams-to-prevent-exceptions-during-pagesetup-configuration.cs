// Title: Safely configure worksheet background image in Aspose.Cells C# with null stream handling and error logging
// AI Prompts: Write a C# method that accepts a workbook file path and an optional image Stream, adds the image as a picture on the first worksheet, and logs warnings if the stream is null or unreadable. | Modify an existing Aspose.Cells page‑setup routine to clear all pictures from a worksheet when no background image stream is provided, ensuring no exceptions are thrown. | Implement file‑existence checks and try‑catch blocks around workbook loading and saving in C# to prevent FileNotFoundException and save errors when using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add a background picture to a worksheet from a Stream safely | C# handle missing background image when setting Excel page setup with Aspose.Cells | clear worksheet pictures in Aspose.Cells if background image stream is null | prevent exceptions during workbook load and save in Aspose.Cells C# example | log warnings instead of throwing when background image file not found Aspose.Cells
// Tags: Aspose.Cells worksheet background image from stream | null stream validation Aspose.Cells | clear worksheet pictures Aspose.Cells | exception handling workbook load save Aspose.Cells C# | page setup error logging Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to load an Excel workbook with Aspose.Cells, optionally set a background picture on the first worksheet using a provided Stream, clear existing pictures when the stream is missing, and safely handle file‑existence, stream readability, and workbook save errors with console logging.
    public class WorkbookPageSetupHelper
    {
        /// <param name="workbookPath">Full path to the Excel file.</param>
        /// <param name="backgroundImageStream">Stream containing the background image (can be null).</param>
        public void ConfigurePageSetup(string workbookPath, Stream? backgroundImageStream)
        {
            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook (lifecycle rule)
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set or clear background image using worksheet pictures (avoids System.Drawing)
            if (backgroundImageStream != null && backgroundImageStream.CanRead)
            {
                try
                {
                    // Ensure the stream is positioned at the beginning
                    if (backgroundImageStream.CanSeek)
                        backgroundImageStream.Position = 0;

                    // Remove existing pictures that might act as background
                    sheet.Pictures.Clear();

                    // Add the image as a picture covering the sheet
                    // The picture is anchored at cell A1 (row 0, column 0)
                    int pictureIndex = sheet.Pictures.Add(0, 0, backgroundImageStream);
                    // Optionally, you could resize the picture to fit the used range or page size here
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set background image. Details: {ex.Message}");
                }
            }
            else
            {
                try
                {
                    // Clear any existing pictures acting as background
                    sheet.Pictures.Clear();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to clear background image. Details: {ex.Message}");
                }
            }

            // Save the workbook (lifecycle rule)
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the workbook as the first argument.");
                return;
            }

            string workbookPath = args[0];
            Stream? backgroundStream = null;

            // Optional background image path
            if (args.Length > 1)
            {
                string imagePath = args[1];
                if (File.Exists(imagePath))
                {
                    try
                    {
                        backgroundStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to open image file. Details: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Image file not found at '{imagePath}'. Continuing without background.");
                }
            }

            var helper = new WorkbookPageSetupHelper();
            helper.ConfigurePageSetup(workbookPath, backgroundStream);

            // Dispose the image stream if it was opened
            backgroundStream?.Dispose();
        }
    }
}
