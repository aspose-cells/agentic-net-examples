using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class SecureExternalPathDemo
{
    static void Main()
    {
        // Define a base directory that is allowed for external files.
        // All external paths will be validated against this directory.
        string baseDir = Path.GetFullPath("AllowedExternalFiles");
        Directory.CreateDirectory(baseDir);

        // Create a new workbook (creation rule).
        Workbook workbook = new Workbook();

        // --------------------------------------------------------------------
        // Example: handling a user‑provided external file path securely.
        // --------------------------------------------------------------------
        // Simulate a path that might come from an external source (e.g., user input).
        string userProvidedPath = @"..\secret\malicious.xlsx";

        // Validate and normalize the path. If it is outside the allowed base directory,
        // GetSafePath returns null and the path is rejected.
        string safePath = GetSafePath(userProvidedPath, baseDir);
        if (safePath != null)
        {
            // If the workbook already contains external connections, set their
            // SourceFile and OdcFile properties using the validated path.
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // Set a safe source file path (property rule).
                conn.SourceFile = safePath;

                // Set a safe ODC file path (property rule). Here we place the ODC file
                // inside the allowed base directory.
                conn.OdcFile = Path.Combine(baseDir, "connection.odc");
            }
        }

        // --------------------------------------------------------------------
        // Example: adding an external link with a validated file name.
        // --------------------------------------------------------------------
        string externalFileName = "ExternalData.xlsx";
        string externalFilePath = Path.Combine(baseDir, externalFileName);

        // Ensure the external workbook exists so that the link is valid.
        if (!File.Exists(externalFilePath))
        {
            new Workbook().Save(externalFilePath);
        }

        // Add the external link using the Add method (method rule).
        // DirectoryType.Volume indicates a full path is supplied.
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
        int linkIndex = externalLinks.Add(DirectoryType.Volume, externalFilePath, new string[] { "Sheet1" });

        // Optionally adjust the link's OriginalDataSource property securely.
        ExternalLink link = externalLinks[linkIndex];
        link.OriginalDataSource = externalFilePath; // property rule

        // --------------------------------------------------------------------
        // Save the workbook to a location inside the allowed base directory.
        // --------------------------------------------------------------------
        string outputPath = Path.Combine(baseDir, "SecureWorkbook.xlsx");
        workbook.Save(outputPath);
    }

    // Helper method that validates a path against a trusted base directory.
    // Returns the absolute path if it is within the base directory; otherwise null.
    static string GetSafePath(string inputPath, string baseDirectory)
    {
        try
        {
            // Resolve to an absolute path.
            string fullPath = Path.GetFullPath(inputPath);

            // Ensure the resolved path starts with the allowed base directory.
            // Comparison is case‑insensitive for Windows file systems.
            if (fullPath.StartsWith(baseDirectory, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath;
            }
            else
            {
                Console.WriteLine($"Rejected unsafe path: {inputPath}");
                return null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing path '{inputPath}': {ex.Message}");
            return null;
        }
    }
}