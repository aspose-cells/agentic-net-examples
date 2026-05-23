using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExample
{
    public class SmartMarkerProcessor
    {
        // Processes a workbook downloaded from a URL, applies smart markers, and saves the result locally.
        public static async Task ProcessWorkbookAsync(string fileUrl, string outputFolderPath, string outputFileName)
        {
            // Ensure the output folder exists.
            if (!Directory.Exists(outputFolderPath))
            {
                Directory.CreateDirectory(outputFolderPath);
            }

            // Full path for the resulting file.
            string outputPath = Path.Combine(outputFolderPath, outputFileName);

            try
            {
                // Inform Aspose.Cells that the code runs in a cloud environment.
                CellsHelper.IsCloudPlatform = true;

                // Download the workbook file into a memory stream.
                using (HttpClient httpClient = new HttpClient())
                using (Stream downloadStream = await httpClient.GetStreamAsync(fileUrl))
                {
                    // Load the workbook from the downloaded stream.
                    Workbook workbook = new Workbook(downloadStream);

                    // Set up the WorkbookDesigner to process smart markers.
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };

                    // OPTIONAL: set a JSON data source if your template uses smart markers.
                    // string json = "{\"Name\":\"John Doe\",\"Value\":123.45}";
                    // designer.SetJsonDataSource("Data", json);

                    // Process all smart markers in the workbook.
                    designer.Process();

                    // Save the processed workbook to a memory stream in XLSX format.
                    using (MemoryStream outStream = new MemoryStream())
                    {
                        workbook.Save(outStream, SaveFormat.Xlsx);
                        outStream.Position = 0; // Reset for reading.

                        // Write the stream to a local file.
                        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                        {
                            await outStream.CopyToAsync(fileStream);
                        }
                    }
                }
            }
            catch (HttpRequestException httpEx)
            {
                Console.Error.WriteLine($"Error downloading the file: {httpEx.Message}");
                throw;
            }
            catch (IOException ioEx)
            {
                Console.Error.WriteLine($"IO error while saving the workbook: {ioEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
                throw;
            }
        }

        // Optional overload to process a local template file.
        public static void ProcessWorkbookFromFile(string templatePath, string outputFolderPath, string outputFileName)
        {
            if (!File.Exists(templatePath))
            {
                Console.Error.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Ensure the output folder exists.
            if (!Directory.Exists(outputFolderPath))
            {
                Directory.CreateDirectory(outputFolderPath);
            }

            string outputPath = Path.Combine(outputFolderPath, outputFileName);

            try
            {
                CellsHelper.IsCloudPlatform = true;
                Workbook workbook = new Workbook(templatePath);
                WorkbookDesigner designer = new WorkbookDesigner { Workbook = workbook };
                designer.Process();
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (IOException ioEx)
            {
                Console.Error.WriteLine($"IO error while processing the workbook: {ioEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        // Entry point for the console application.
        public static async Task Main(string[] args)
        {
            try
            {
                if (args.Length == 3)
                {
                    // args: fileUrl outputFolderPath outputFileName
                    await SmartMarkerProcessor.ProcessWorkbookAsync(args[0], args[1], args[2]);
                }
                else if (args.Length == 3 && File.Exists(args[0]))
                {
                    // args: templatePath outputFolderPath outputFileName
                    SmartMarkerProcessor.ProcessWorkbookFromFile(args[0], args[1], args[2]);
                }
                else
                {
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  dotnet run <fileUrl> <outputFolderPath> <outputFileName>");
                    Console.WriteLine("  or");
                    Console.WriteLine("  dotnet run <templatePath> <outputFolderPath> <outputFileName>");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}