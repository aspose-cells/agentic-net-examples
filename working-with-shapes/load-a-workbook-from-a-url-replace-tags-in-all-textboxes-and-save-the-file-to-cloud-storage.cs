using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

public class WorkbookProcessor
{
    // Entry point
    public static async Task Main(string[] args)
    {
        try
        {
            // Example usage
            string fileUrl = "https://example.com/sample.xlsx";          // URL of the source workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "sample_processed.xlsx"); // Local output path

            await ProcessWorkbookAsync(fileUrl, outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Loads a workbook from a URL, replaces placeholder tags in all TextBoxes,
    /// and saves the resulting file to the specified path.
    /// </summary>
    private static async Task ProcessWorkbookAsync(string fileUrl, string outputFilePath)
    {
        // Ensure the output directory exists
        string? directory = Path.GetDirectoryName(outputFilePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // 1. Download the Excel file as a stream.
        using var httpClient = new HttpClient();
        await using Stream downloadStream = await httpClient.GetStreamAsync(fileUrl);

        // 2. Load the workbook from the downloaded stream.
        var workbook = new Workbook(downloadStream);

        // 3. Iterate through all worksheets and their TextBoxes to replace tags.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Aspose.Cells.Drawing.TextBox textBox in sheet.TextBoxes)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    // Replace placeholder with actual value.
                    textBox.Text = textBox.Text.Replace("{{PLACEHOLDER}}", "ActualValue");
                }
            }
        }

        // 4. Save the modified workbook to the specified file.
        workbook.Save(outputFilePath, SaveFormat.Xlsx);

        // Clean up.
        workbook.Dispose();
    }
}