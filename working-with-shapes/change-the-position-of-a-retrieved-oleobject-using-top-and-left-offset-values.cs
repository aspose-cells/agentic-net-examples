// Title: Move an OleObject in Aspose.Cells for .NET – adjust Top and Left pixel offsets
// Description: This example creates a workbook, inserts an OleObject at row 5, column 3 (150 × 150 px), optionally sets a custom icon, then shifts the object 20 px down and 30 px right by modifying its Top and Left properties before saving the file.
// Keywords: Aspose.Cells OleObject reposition | C# move OleObject top left | adjust OleObject pixel offset | Aspose.Cells set object coordinates | change embedded OLE object location
// Common Searches: how to change OleObject position Aspose.Cells | move embedded OLE object by pixels C# | adjust top left of OleObject in Excel using Aspose | Aspose.Cells offset OleObject coordinates
// Developer Intent: Shift a retrieved OleObject by adding specific pixel values to its Top and Left properties.
// Use Cases: Prevent overlap with other shapes after inserting an OleObject. | Apply a uniform offset to a series of OleObjects for consistent alignment. | Programmatically place OleObjects based on runtime calculations such as centering within a cell range.
// AI Prompts: Generate C# code that moves an OleObject a given number of pixels using Aspose.Cells. | Explain how to retrieve an OleObject from a worksheet and modify its Top and Left values with offsets. | Show how to calculate pixel offsets for precise OleObject placement relative to Excel cells in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectPositionExample
{
    // This example creates a workbook, inserts an OleObject at row 5, column 3 (150 × 150 px), optionally sets a custom icon, then shifts the object 20 px down and 30 px right by modifying its Top and Left properties before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to custom icon image (optional)
                string iconPath = "sample_icon.png";

                // Add an OleObject at row 5, column 3 with size 150x150 pixels
                // Use an empty byte array as placeholder OLE data
                int oleIndex = worksheet.OleObjects.Add(5, 3, 150, 150, new byte[0]);

                // Retrieve the OleObject that was just added
                OleObject ole = worksheet.OleObjects[oleIndex];

                // If the icon file exists, attempt to assign it (property may not be available in older versions)
                if (File.Exists(iconPath))
                {
                    try
                    {
                        // The IconFilePath property is available in recent Aspose.Cells versions.
                        // If not present, this block will be skipped without affecting execution.
                        ole.GetType().GetProperty("IconFilePath")?.SetValue(ole, iconPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to set icon file: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Icon file not found: {iconPath}. Using default icon.");
                }

                // Define offset values (in pixels)
                int topOffset = 20;   // move down by 20 pixels
                int leftOffset = 30;  // move right by 30 pixels

                // Adjust the OleObject's position
                ole.Top += topOffset;
                ole.Left += leftOffset;

                // Display the new position in the console
                Console.WriteLine($"OleObject new Top: {ole.Top} pixels");
                Console.WriteLine($"OleObject new Left: {ole.Left} pixels");

                // Save the workbook to a file
                string outputPath = "OleObjectMoved.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
