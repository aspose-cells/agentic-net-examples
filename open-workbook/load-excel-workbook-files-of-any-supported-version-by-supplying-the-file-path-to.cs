using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookDemo
    {
        // Loads a workbook from the specified file path using the default constructor.
        public static void Run(string filePath)
        {
            // Initialize the Workbook by providing the file path.
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet to demonstrate that the file is loaded.
            Worksheet firstSheet = workbook.Worksheets[0];

            // Output basic information about the loaded workbook.
            Console.WriteLine($"Workbook loaded from: {filePath}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");
        }

        // Loads a workbook from the specified file path using custom LoadOptions.
        public static void RunWithOptions(string filePath)
        {
            // Create LoadOptions (you can customize properties here if needed).
            LoadOptions loadOptions = new LoadOptions();

            // Initialize the Workbook with the file path and the LoadOptions instance.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Output the detected file format to verify successful loading.
            Console.WriteLine($"Workbook loaded with options from: {filePath}");
            Console.WriteLine($"Detected file format: {workbook.FileFormat}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as a command‑line argument.");
                return;
            }

            string filePath = args[0];

            try
            {
                LoadWorkbookDemo.Run(filePath);
                LoadWorkbookDemo.RunWithOptions(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}