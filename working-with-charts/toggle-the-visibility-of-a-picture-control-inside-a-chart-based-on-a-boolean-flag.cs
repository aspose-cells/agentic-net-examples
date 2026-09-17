// Title: How to toggle the visibility of a named picture inside an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, locates a chart picture called "MyPicture", and updates its IsVisible flag according to a Boolean variable. | Show how to employ .NET reflection to retrieve the hidden Pictures collection of a Chart object in Aspose.Cells and change a picture's visibility. | Create a snippet that conditionally hides or shows a chart image and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# hide chart picture called MyPicture based on a condition | how to change chart image visibility programmatically with Aspose.Cells | reflection technique for accessing chart pictures in Aspose.Cells .NET | conditional display of Excel chart picture using Aspose.Cells C#
// Tags: Aspose.Cells chart picture visibility control | C# reflection Aspose.Cells chart pictures | Excel chart picture conditional visibility | modify chart picture IsVisible Aspose | programmatic chart picture toggle C#

using System;
using System.IO;
using System.Collections;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, accesses the first chart, uses reflection to enumerate its Pictures collection, finds the picture named "MyPicture", sets its IsVisible property based on a Boolean flag, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first chart on the worksheet (adjust index or name as needed)
            Chart chart = worksheet.Charts[0];

            // Flag that determines visibility of the picture
            bool showPicture = true; // set to false to hide the picture

            // Attempt to manipulate chart pictures via reflection to stay compatible with older versions
            try
            {
                PropertyInfo picturesProp = typeof(Chart).GetProperty("Pictures");
                if (picturesProp != null)
                {
                    object picturesObj = picturesProp.GetValue(chart);
                    if (picturesObj is IEnumerable pictures)
                    {
                        foreach (object picture in pictures)
                        {
                            Type picType = picture.GetType();
                            PropertyInfo nameProp = picType.GetProperty("Name");
                            PropertyInfo visibleProp = picType.GetProperty("IsVisible");
                            if (nameProp != null && visibleProp != null)
                            {
                                string name = nameProp.GetValue(picture) as string;
                                if (name == "MyPicture")
                                {
                                    visibleProp.SetValue(picture, showPicture);
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Chart picture manipulation is not supported in the current Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing chart pictures: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
