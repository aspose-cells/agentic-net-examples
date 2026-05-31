using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsSecurityDemo
{
    public static class ExternalPathHandler
    {
        // Base directory that is allowed for external file references
        private static readonly string AllowedBaseDirectory = Path.GetFullPath(@"C:\AllowedData");

        // Directory where the processed workbook will be saved
        private static readonly string OutputDirectory = Path.GetFullPath(@"C:\ProcessedWorkbooks");

        public static void Run(string inputWorkbookPath)
        {
            // Verify input file exists
            if (!File.Exists(inputWorkbookPath))
                throw new FileNotFoundException("Input workbook not found.", inputWorkbookPath);

            // Resolve the full path of the input workbook and ensure it is within the allowed base directory
            string fullInputPath = Path.GetFullPath(inputWorkbookPath);
            if (!IsPathWithinBase(fullInputPath, AllowedBaseDirectory))
                throw new UnauthorizedAccessException("Input workbook path is outside the allowed directory.");

            Workbook workbook;
            try
            {
                // Load the workbook (standard Aspose.Cells lifecycle)
                workbook = new Workbook(fullInputPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to load workbook.", ex);
            }

            // Iterate through all external connections and sanitize their SourceFile paths
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                string sourceFile = connection.SourceFile;
                if (string.IsNullOrEmpty(sourceFile))
                    continue;

                // Convert to absolute path (handles both URI and system-specific notation)
                string absoluteSourcePath;
                try
                {
                    // If the source is a URI, attempt to get the local path; otherwise treat as file path
                    if (Uri.IsWellFormedUriString(sourceFile, UriKind.Absolute))
                    {
                        Uri uri = new Uri(sourceFile);
                        absoluteSourcePath = uri.IsFile ? uri.LocalPath : string.Empty;
                    }
                    else
                    {
                        absoluteSourcePath = Path.GetFullPath(sourceFile);
                    }
                }
                catch
                {
                    // If parsing fails, skip this connection for safety
                    continue;
                }

                // Verify the resolved path is within the allowed base directory
                if (IsPathWithinBase(absoluteSourcePath, AllowedBaseDirectory))
                {
                    // Keep the safe path
                    connection.SourceFile = absoluteSourcePath;
                }
                else
                {
                    // Clear the path if it is unsafe
                    connection.SourceFile = string.Empty;
                }
            }

            // Ensure the output directory exists
            if (!Directory.Exists(OutputDirectory))
                Directory.CreateDirectory(OutputDirectory);

            // Construct a safe output file name (avoid path traversal by using only the file name)
            string safeFileName = Path.GetFileName(fullInputPath);
            string outputPath = Path.Combine(OutputDirectory, safeFileName);

            // Save the workbook (standard Aspose.Cells save lifecycle)
            workbook.Save(outputPath);
        }

        // Helper method to verify that a path is inside a given base directory
        private static bool IsPathWithinBase(string path, string baseDir)
        {
            // Normalize both paths
            string normalizedPath = Path.GetFullPath(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            string normalizedBase = Path.GetFullPath(baseDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

            // Compare using case-insensitive comparison on Windows
            return normalizedPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string inputPath;
                if (args.Length > 0)
                {
                    inputPath = args[0];
                }
                else
                {
                    Console.WriteLine("Please provide the full path to the input workbook as a command‑line argument.");
                    return;
                }

                ExternalPathHandler.Run(inputPath);
                Console.WriteLine("Workbook processed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}