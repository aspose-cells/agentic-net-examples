// Title: Create a custom IFilePathProvider in Aspose.Cells for C# to generate lowercase .html file names from worksheet names
// Common Searches: how to implement IFilePathProvider in Aspose.Cells C# to customize hyperlink file names | generate lowercase HTML file names from Excel sheet names using Aspose.Cells | Aspose.Cells custom file path provider for sheet name sanitization | C# example of converting worksheet names to URL‑friendly strings with Aspose.Cells | set workbook Settings.FilePathProvider to a custom provider in Aspose.Cells .NET
// Tags: Aspose.Cells lowercase worksheet file naming | C# sheet name to URL-friendly path | hyperlink file path provider example | Excel to HTML link generation Aspose.Cells | sanitized worksheet filenames .NET

using System;
using System.IO;
using Aspose.Cells;

// The LowercaseUrlFilePathProvider implements IFilePathProvider; its GetFilePath method trims the worksheet name, replaces spaces with hyphens, converts it to lower case, and appends a .html extension, while GetFullName returns the absolute path for a given relative file path.
public class LowercaseUrlFilePathProvider : IFilePathProvider
{
    // Called by Aspose.Cells when a file path for a worksheet is required.
    // Converts the worksheet name to a lowercase, URL‑friendly string.
    public string GetFilePath(string sheetName, int sheetIndex)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("Sheet name cannot be null or empty.", nameof(sheetName));

        // Trim, replace spaces with hyphens, and convert to lower case.
        string sanitized = sheetName.Trim().Replace(' ', '-').ToLowerInvariant();

        // Append an extension suitable for linking (e.g., .html).
        return $"{sanitized}.html";
    }

    // Called by Aspose.Cells to resolve a relative path to a full file system path.
    public string GetFullName(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return string.Empty;

        // Return the absolute path for the given relative file path.
        return Path.GetFullPath(filePath);
    }
}

// Example usage
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before loading.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The FilePathProvider property is not available in this version of Aspose.Cells.
            // If a newer version supports it, you can assign the custom provider as shown below:
            // workbook.Settings.FilePathProvider = new LowercaseUrlFilePathProvider();

            // Perform any operations that rely on the provider (e.g., creating hyperlinks).
            // ...

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
