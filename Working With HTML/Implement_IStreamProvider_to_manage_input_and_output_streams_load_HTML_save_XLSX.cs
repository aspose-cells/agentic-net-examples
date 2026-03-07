using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that supplies a FileStream for reading resources
    public class CustomStreamProvider : IStreamProvider
    {
        // Called by Aspose.Cells when a stream is required
        public void InitStream(StreamProviderOptions options)
        {
            // Use the default path supplied by the library and open it for reading
            // Mark the loading type as user provided so the library uses this stream
            options.ResourceLoadingType = ResourceLoadingType.UserProvided;
            options.Stream = File.OpenRead(options.DefaultPath);
        }

        // Called after the library finishes using the stream
        public void CloseStream(StreamProviderOptions options)
        {
            // Safely close and dispose the stream if it exists
            options.Stream?.Close();
            options.Stream = null;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source HTML file that will be loaded
            string htmlInputPath = "input.html";

            // Path where the resulting XLSX file will be saved
            string xlsxOutputPath = "output.xlsx";

            // Ensure the input file exists for the demo (in real scenarios the file should already exist)
            if (!File.Exists(htmlInputPath))
            {
                // Create a simple HTML file with a table for demonstration purposes
                File.WriteAllText(htmlInputPath,
                    "<html><body><table><tr><td>A1</td><td>Value1</td></tr><tr><td>A2</td><td>Value2</td></tr></table></body></html>");
            }

            // Configure load options to use the custom stream provider
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new CustomStreamProvider()
            };

            // Load the HTML workbook using the custom provider
            Workbook workbook = new Workbook(htmlInputPath, loadOptions);

            // Optionally manipulate the workbook (e.g., add a new sheet)
            Worksheet newSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            newSheet.Cells["A1"].PutValue("Added by StreamProvider demo");

            // Save the workbook to XLSX format
            workbook.Save(xlsxOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlInputPath}' loaded and saved as XLSX to '{xlsxOutputPath}'.");
        }
    }
}