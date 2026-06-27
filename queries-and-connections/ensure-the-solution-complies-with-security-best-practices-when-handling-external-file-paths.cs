using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class SecureExternalPathHandling
{
    static void Main()
    {
        try
        {
            // Define a base directory that is considered safe for all external files.
            string baseDirectory = Path.GetFullPath("AllowedExternalFiles");
            Directory.CreateDirectory(baseDirectory);

            // Example input workbook path located inside the safe directory.
            string inputPath = Path.Combine(baseDirectory, "input.xlsx");

            // Verify the input path does not escape the allowed directory.
            if (!IsPathSecure(inputPath, baseDirectory))
            {
                Console.WriteLine("Input path is not within the allowed directory.");
                return;
            }

            Workbook workbook;

            // Load the workbook only if the file exists; otherwise create a new workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file not found at {inputPath}. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook
            }

            // Iterate over external connections and sanitize their file references.
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // Secure handling of SourceFile property.
                if (!string.IsNullOrEmpty(conn.SourceFile))
                {
                    string secureSource = GetSecurePath(conn.SourceFile, baseDirectory);
                    if (secureSource != null)
                    {
                        conn.SourceFile = secureSource;
                    }
                }

                // Secure handling of OdcFile property.
                if (!string.IsNullOrEmpty(conn.OdcFile))
                {
                    string secureOdc = GetSecurePath(conn.OdcFile, baseDirectory);
                    if (secureOdc != null)
                    {
                        conn.OdcFile = secureOdc;
                    }
                }
            }

            // Add an external link only if the target file resides within the safe directory.
            string externalFile = Path.Combine(baseDirectory, "ExternalWorkbook.xlsx");
            if (File.Exists(externalFile) && IsPathSecure(externalFile, baseDirectory))
            {
                string[] sheetNames = new string[] { "Sheet1", "Sheet2" };
                int linkIndex = workbook.Worksheets.ExternalLinks.Add(DirectoryType.Volume, externalFile, sheetNames);
                Console.WriteLine($"External link added at index {linkIndex}");
            }

            // Save the workbook to a location inside the allowed directory (lifecycle rule).
            string outputPath = Path.Combine(baseDirectory, "output_secure.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Returns a normalized absolute path if it stays within baseDir; otherwise null.
    static string GetSecurePath(string path, string baseDir)
    {
        try
        {
            // Resolve relative paths against the base directory.
            string combined = Path.IsPathRooted(path) ? path : Path.Combine(baseDir, path);
            string fullPath = Path.GetFullPath(combined);

            // Ensure the resolved path starts with the allowed base directory.
            if (fullPath.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath;
            }
        }
        catch
        {
            // Ignore malformed paths.
        }

        Console.WriteLine($"Rejected insecure path: {path}");
        return null;
    }

    // Simple validation that a given path resolves inside the allowed base directory.
    static bool IsPathSecure(string path, string baseDir)
    {
        try
        {
            string fullPath = Path.GetFullPath(path);
            return fullPath.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }
}