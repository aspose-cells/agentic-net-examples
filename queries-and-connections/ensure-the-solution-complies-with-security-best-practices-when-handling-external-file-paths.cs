// Title: Secure External File Handling in Aspose.Cells: Load, Validate, and Link Workbooks within a Trusted Directory (C#)
// Description: Shows how to create a trusted folder, load a workbook only from that location, sanitize existing external connections, and add a new external link safely using Aspose.Cells for .NET. Includes path validation, placeholder workbook creation, and exception handling to block path‑traversal attacks.
// Keywords: Aspose.Cells | C# | secure file paths | external connections | trusted directory | path traversal protection | workbook loading | external links | data validation | .NET security
// Common Searches: Aspose.Cells restrict workbook loading to specific folder | sanitize external data connections in Aspose.Cells | add external link safely with Aspose.Cells .NET | prevent path traversal when using Aspose.Cells | validate file path before loading workbook Aspose.Cells
// Developer Intent: Load a workbook only from a predefined safe directory, clean any insecure external connections, and add a new external link that also resides within that directory.
// Use Cases: Enforce a corporate policy that all Excel files processed by Aspose.Cells must live in a controlled folder. | Automatically remove or neutralize external data sources that point outside the allowed path. | Create placeholder workbooks for missing external files and link them with absolute, validated paths. | Throw clear security exceptions when a requested file or link is outside the trusted area.
// AI Prompts: Write a logger that records every insecure external connection removed from an Aspose.Cells workbook. | Generate unit tests for IsPathInBaseDirectory covering normal, edge, and traversal cases. | Provide an alternative implementation that uses Path.GetRelativePath to enforce trusted‑directory constraints when adding external links.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Shows how to create a trusted folder, load a workbook only from that location, sanitize existing external connections, and add a new external link safely using Aspose.Cells for .NET. Includes path validation, placeholder workbook creation, and exception handling to block path‑traversal attacks.
class SecureExternalPathHandler
{
    static void Main()
    {
        try
        {
            // Define a trusted base directory for all external files.
            string trustedBaseDir = Path.GetFullPath(@"C:\TrustedData");
            Directory.CreateDirectory(trustedBaseDir);

            // Path to the workbook that will be processed (must reside in the trusted directory).
            string workbookPath = Path.Combine(trustedBaseDir, "input.xlsx");

            // Ensure the workbook exists; create a minimal one if it does not.
            if (!File.Exists(workbookPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(workbookPath);
                Console.WriteLine($"Created placeholder workbook at {workbookPath}");
            }

            // Load the workbook using a secure helper that validates the location.
            Workbook workbook = LoadWorkbookSecurely(workbookPath, trustedBaseDir);

            // Validate and sanitize any external connections defined in the workbook.
            SecureExternalConnections(workbook, trustedBaseDir);

            // Add a new external link, ensuring the target file is also inside the trusted directory.
            AddSecureExternalLink(workbook, trustedBaseDir, "ExternalData.xlsx", new[] { "Sheet1", "Sheet2" });

            // Save the modified workbook back to the trusted location.
            string outputPath = Path.Combine(trustedBaseDir, "output.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved securely to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Loads a workbook only if the file resides within the allowed base directory.
    static Workbook LoadWorkbookSecurely(string path, string baseDir)
    {
        string fullPath = Path.GetFullPath(path);

        if (!IsPathInBaseDirectory(fullPath, baseDir))
            throw new UnauthorizedAccessException("Attempt to load workbook from outside the trusted directory.");

        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"Workbook not found at {fullPath}");

        return new Workbook(fullPath);
    }

    // Iterates over external connections and removes or normalizes any paths that are not trusted.
    static void SecureExternalConnections(Workbook workbook, string baseDir)
    {
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            if (!string.IsNullOrEmpty(conn.SourceFile))
            {
                string sourceFullPath = Path.GetFullPath(conn.SourceFile);
                if (IsPathInBaseDirectory(sourceFullPath, baseDir))
                {
                    // Normalize to a full absolute path.
                    conn.SourceFile = sourceFullPath;
                }
                else
                {
                    // Detected an insecure path – clear it and log the incident.
                    Console.WriteLine($"Insecure external source detected and removed: {conn.SourceFile}");
                    conn.SourceFile = string.Empty;
                }
            }
        }
    }

    // Adds an external link after confirming the target file is within the trusted directory.
    static void AddSecureExternalLink(Workbook workbook, string baseDir, string fileName, string[] sheetNames)
    {
        string externalFullPath = Path.GetFullPath(Path.Combine(baseDir, fileName));

        if (!IsPathInBaseDirectory(externalFullPath, baseDir))
            throw new UnauthorizedAccessException("External link points outside the trusted directory.");

        // Ensure the external file exists; create a placeholder if necessary.
        if (!File.Exists(externalFullPath))
        {
            var extWb = new Workbook();
            foreach (string sheet in sheetNames)
                extWb.Worksheets.Add(sheet);
            extWb.Save(externalFullPath);
            Console.WriteLine($"Created placeholder external workbook at {externalFullPath}");
        }

        // Add the external link using an absolute path.
        int linkIndex = workbook.Worksheets.ExternalLinks.Add(DirectoryType.Volume, externalFullPath, sheetNames);
        ExternalLink link = workbook.Worksheets.ExternalLinks[linkIndex];
        Console.WriteLine($"Added external link to {link.DataSource} referencing {sheetNames.Length} sheet(s).");
    }

    // Helper that checks whether a given path is a descendant of the base directory.
    static bool IsPathInBaseDirectory(string path, string baseDir)
    {
        string normalizedBase = baseDir.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string normalizedPath = Path.GetFullPath(path);
        return normalizedPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase);
    }
}
