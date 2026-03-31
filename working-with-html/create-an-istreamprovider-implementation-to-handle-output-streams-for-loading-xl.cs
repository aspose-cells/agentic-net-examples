using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom implementation of IStreamProvider to manage streams for both saving and loading.
    public class FileStreamProvider : IStreamProvider
    {
        private readonly string _baseDirectory;

        public FileStreamProvider(string baseDirectory)
        {
            _baseDirectory = baseDirectory;
        }

        // Called by Aspose.Cells when a stream is required.
        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the base directory exists.
            Directory.CreateDirectory(_baseDirectory);

            // Determine the target file name.
            // If DefaultPath is provided (used during loading), use it; otherwise create a name based on the operation.
            string fileName = options.DefaultPath != null
                ? Path.GetFileName(options.DefaultPath)
                : Guid.NewGuid().ToString() + ".tmp";

            string fullPath = Path.Combine(_baseDirectory, fileName);

            // Set the custom path that will be written into the generated HTML (if applicable).
            options.CustomPath = fullPath;

            // Create the file stream for writing (or reading if the file already exists).
            // For loading external resources Aspose may request a read stream; we handle both cases.
            FileMode mode = File.Exists(fullPath) ? FileMode.Open : FileMode.Create;
            FileAccess access = File.Exists(fullPath) ? FileAccess.ReadWrite : FileAccess.Write;

            options.Stream = new FileStream(fullPath, mode, access);
        }

        // Called by Aspose.Cells after the operation is finished.
        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Directory where all generated files will be placed.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeStreamProviderDemo");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a workbook and save it as HTML using the custom provider.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello from Aspose.Cells!");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.StreamProvider = new FileStreamProvider(outputDir);

            string htmlPath = Path.Combine(outputDir, "sample.html");
            workbook.Save(htmlPath, saveOptions);
            Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

            // -----------------------------------------------------------------
            // 2. Load the previously saved HTML back into a workbook using the same provider.
            // -----------------------------------------------------------------
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.StreamProvider = new FileStreamProvider(outputDir);

            Workbook loadedWorkbook = new Workbook(htmlPath, loadOptions);
            string loadedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded cell A1 value: {loadedValue}");
        }
    }
}