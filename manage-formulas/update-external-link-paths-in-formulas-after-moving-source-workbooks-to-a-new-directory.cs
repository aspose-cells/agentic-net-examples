// Title: C# – Update external link paths in Excel formulas after moving source workbooks with Aspose.Cells
// Common Searches: aspnet cells change external link source directory in an Excel workbook | c# replace old folder path for external references using Aspose.Cells | how to adjust ExternalLinks SourceFullName after moving source files in .NET | update workbook external links when source files are relocated with Aspose.Cells | dynamic binding to modify external link paths in Excel via Aspose.Cells C#
// Tags: update external links Aspose.Cells | modify SourceFullName property | dynamic binding ExternalLinks API | handle missing ExternalLinks support | save workbook after link path change | relocate source workbooks Excel formulas

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;
using Microsoft.CSharp.RuntimeBinder;

// The example loads a target workbook, defines the previous and new source directories, and iterates through each worksheet. Using dynamic binding it accesses the ExternalLinks collection and the SourceFullName of each link. When a link path starts with the old directory, the code builds a new path based on the new directory and assigns it back. The routine safely skips sheets or links if the ExternalLinks API is unavailable, then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the target workbook (the one containing external links)
            string targetPath = @"C:\Data\TargetWorkbook.xlsx";
            string updatedPath = @"C:\Data\TargetWorkbook_Updated.xlsx";

            // Verify that the target workbook exists
            if (!File.Exists(targetPath))
                throw new FileNotFoundException("Target workbook not found.", targetPath);

            // Load the workbook
            var workbook = new Workbook(targetPath);

            // Old and new directories for the source workbooks
            string oldDirectory = @"C:\OldSourceDir\";
            string newDirectory = @"D:\NewSourceDir\";

            // Ensure both directories end with a separator character
            if (!oldDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()))
                oldDirectory += Path.DirectorySeparatorChar;
            if (!newDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()))
                newDirectory += Path.DirectorySeparatorChar;

            // Attempt to update external link paths using dynamic binding (works only if the API exists at runtime)
            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Use dynamic to avoid compile‑time binding to possibly missing members
                    dynamic dynSheet = sheet;
                    IEnumerable externalLinks = null;

                    try
                    {
                        externalLinks = dynSheet.ExternalLinks as IEnumerable;
                    }
                    catch (RuntimeBinderException)
                    {
                        // ExternalLinks property not available in this version; skip this sheet
                        continue;
                    }

                    if (externalLinks == null) continue;

                    foreach (var link in externalLinks)
                    {
                        dynamic dynLink = link;
                        string currentPath = null;

                        try
                        {
                            currentPath = dynLink.SourceFullName as string;
                        }
                        catch (RuntimeBinderException)
                        {
                            // SourceFullName property not available; skip this link
                            continue;
                        }

                        if (!string.IsNullOrEmpty(currentPath) &&
                            currentPath.StartsWith(oldDirectory, StringComparison.OrdinalIgnoreCase))
                        {
                            string relativePart = currentPath.Substring(oldDirectory.Length);
                            string updatedLinkPath = Path.Combine(newDirectory, relativePart);

                            try
                            {
                                dynLink.SourceFullName = updatedLinkPath;
                            }
                            catch (RuntimeBinderException)
                            {
                                // Unable to set SourceFullName; ignore and continue
                            }
                        }
                    }
                }
            }
            catch (Exception ex) when (ex is MissingMethodException || ex is RuntimeBinderException)
            {
                // ExternalLinks API not supported; skip link updating.
                Console.WriteLine("ExternalLinks API not supported in this Aspose.Cells version. Skipping link update.");
            }

            // Save the workbook with updated external link paths (or unchanged if skipped)
            workbook.Save(updatedPath);
            Console.WriteLine($"Workbook saved successfully to '{updatedPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
