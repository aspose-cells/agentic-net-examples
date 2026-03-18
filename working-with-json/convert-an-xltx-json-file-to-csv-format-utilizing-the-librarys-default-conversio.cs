using System;
using Aspose.Cells;

class ConvertXltxJsonToCsv
{
    static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Path to the source XLTX file (Excel template)
        string sourcePath = "input.xltx";

        // Desired CSV output path
        string destinationPath = "output.csv";

        try
        {
            // Load the workbook from the XLTX file
            var workbook = new Workbook(sourcePath);

            // Save the workbook as CSV using TxtSaveOptions
            var saveOptions = new TxtSaveOptions(SaveFormat.CSV);
            saveOptions.Separator = ',';
            workbook.Save(destinationPath, saveOptions);

            Console.WriteLine($"Conversion succeeded: '{sourcePath}' -> '{destinationPath}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}