// Title: Securely validate file paths when loading and saving Excel workbooks with Aspose.Cells in C#
// Common Searches: how to prevent path traversal attacks while using Aspose.Cells in a C# application | C# code to validate Excel file paths before loading with Aspose.Cells | restrict Aspose.Cells workbook save to a specific safe directory in .NET | validate absolute path against a whitelist folder when processing Excel files in C# | example of secure FileStream usage for Aspose.Cells workbook read/write
// Tags: Aspose.Cells directory confinement | C# absolute path validation | restricted folder workbook save Aspose.Cells | safe base folder check before load | Excel workbook FileStream security

using System;
using System.IO;
using Aspose.Cells;

// The sample defines fixed input and output Excel file paths, validates each with a custom IsPathSecure method that checks for null or empty values, invalid characters, absolute paths, resolves the full path, and ensures the location resides within a predefined safe base directory. It then loads the workbook from the input path if it exists (using a read‑only FileStream) or creates a new workbook, adds a worksheet, creates the output directory if needed, and saves the workbook to the output path via a write‑only FileStream, handling any exceptions.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\SafeFolder\input.xlsx";
        string outputPath = @"C:\SafeFolder\output.xlsx";

        // Validate the file paths to prevent path traversal and other attacks
        if (!IsPathSecure(inputPath) || !IsPathSecure(outputPath))
        {
            Console.WriteLine("One or more file paths are invalid or insecure.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new Workbook(fs);
                }
            }
            else
            {
                workbook = new Workbook(); // create empty workbook
                Console.WriteLine("Input file not found. A new workbook was created.");
            }

            // Example operation: add a new worksheet
            workbook.Worksheets.Add("NewSheet");

            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the workbook to the validated output path using a FileStream
            using (FileStream fsOut = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                workbook.Save(fsOut, SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook processed and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }

    // Security check for file paths
    static bool IsPathSecure(string path)
    {
        // Null or empty check
        if (string.IsNullOrWhiteSpace(path))
            return false;

        // Reject paths containing invalid characters
        char[] invalidChars = Path.GetInvalidPathChars();
        if (path.IndexOfAny(invalidChars) >= 0)
            return false;

        // Must be an absolute path
        if (!Path.IsPathRooted(path))
            return false;

        // Resolve to a full path to eliminate ".." segments
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(path);
        }
        catch
        {
            return false;
        }

        // Define a safe base directory (adjust as needed)
        string safeBase = Path.GetFullPath(@"C:\SafeFolder\");

        // Ensure the resolved path resides within the safe base directory
        return fullPath.StartsWith(safeBase, StringComparison.OrdinalIgnoreCase);
    }
}
