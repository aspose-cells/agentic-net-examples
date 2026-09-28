// Title: Export a chart from an Excel workbook to a JPEG file using Aspose.Cells ToImage with reflection in C#
// AI Prompts: Generate C# code that loads an .xlsx file, retrieves the first worksheet chart, and saves it as a JPEG using Aspose.Cells Chart.ToImage invoked via reflection. | Show how to adapt the reflection logic to export the same chart as a PNG image instead of JPEG. | Create a reusable C# method that accepts input workbook path, output image path, and desired image format, then returns a success flag after exporting the chart with Aspose.Cells.
// Common Searches: how to use Aspose.Cells ToImage method with reflection to save a chart as jpg | export Excel chart to image in C# without System.Drawing.Common | save first worksheet chart to jpeg using Aspose.Cells API | C# code example for converting workbook chart to jpg file | invoke Chart.ToImage via reflection for chart image export
// Tags: chart export to jpeg via Aspose.Cells | Aspose.Cells ToImage reflection usage | C# chart image generation without System.Drawing | Excel chart to jpg conversion Aspose.Cells | invoke Chart.ToImage with ImageFormat

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;
using System.Reflection;

// Loads an Excel workbook, obtains the first chart from the first worksheet, and uses reflection to call Chart.ToImage with a JPEG ImageFormat, writing the resulting image to a .jpg file.
class ExportChart
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.jpg";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if necessary)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Load ImageFormat type via reflection to avoid direct dependency on System.Drawing.Common
            Type imageFormatType = Type.GetType("System.Drawing.Imaging.ImageFormat, System.Drawing.Common");
            if (imageFormatType == null)
            {
                Console.WriteLine("Error: Unable to load ImageFormat type. Ensure System.Drawing.Common is available.");
                return;
            }

            // Get the static JPEG format instance
            object jpegFormat = imageFormatType.GetProperty("Jpeg")?.GetValue(null);
            if (jpegFormat == null)
            {
                Console.WriteLine("Error: Unable to obtain JPEG ImageFormat.");
                return;
            }

            // Find the ToImage method that accepts (Stream, ImageFormat)
            MethodInfo toImageMethod = chart.GetType().GetMethod("ToImage", new[] { typeof(Stream), imageFormatType });
            if (toImageMethod == null)
            {
                Console.WriteLine("Error: Unable to locate the ToImage method on the Chart object.");
                return;
            }

            // Export the chart to a JPEG image file using the reflected method
            using (FileStream imageStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                toImageMethod.Invoke(chart, new object[] { imageStream, jpegFormat });
            }

            Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
