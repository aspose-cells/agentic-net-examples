// Title: Prompt user for workbook and new link paths, then update the first external link and save a backup using Aspose.Cells for .NET
// Common Searches: how to update the first external link in an Excel workbook with Aspose.Cells C# console application | Aspose.Cells C# prompt user for new external link file path before changing the link | replace external link source and create a backup copy of the workbook using Aspose.Cells .NET | using dynamic ExternalLinks collection to modify external links in a .xlsx file with Aspose.Cells | save backup workbook after changing external link with Aspose.Cells C#
// Tags: Aspose.Cells update external link source | C# dynamic ExternalLinks collection | Excel workbook backup save Aspose.Cells | console prompt for file path C# | replace first external link .xlsx Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsRibbonDemo
{
    // Console application that updates the first external link in a workbook.
    // A C# console program that asks the user to enter the paths of an Excel workbook and a new external link file, loads the workbook with Aspose.Cells, accesses the first external link via the dynamic ExternalLinks collection, changes its source to the supplied path, and writes a backup copy of the workbook.
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Console.WriteLine("=== Aspose.Cells External Link Updater ===");

                // Prompt for the workbook path
                Console.Write("Enter the full path of the workbook (e.g., C:\\Temp\\SampleWithExternalLinks.xlsx): ");
                string workbookPath = Console.ReadLine()?.Trim('"');

                if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
                {
                    Console.WriteLine("Error: Workbook file not found.");
                    return;
                }

                // Prompt for the new external link file path
                Console.Write("Enter the full path of the new external link file: ");
                string newLinkPath = Console.ReadLine()?.Trim('"');

                if (string.IsNullOrWhiteSpace(newLinkPath) || !File.Exists(newLinkPath))
                {
                    Console.WriteLine("Error: New external link file not found.");
                    return;
                }

                // Load the workbook
                Workbook wb;
                try
                {
                    wb = new Workbook(workbookPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }

                // Use dynamic to access ExternalLinks (may not exist in older versions)
                dynamic worksheet = wb.Worksheets[0];
                dynamic externalLinks;
                try
                {
                    externalLinks = worksheet.ExternalLinks;
                }
                catch (RuntimeBinderException)
                {
                    Console.WriteLine("External links are not supported in the loaded Aspose.Cells version.");
                    return;
                }

                // Ensure we have a collection with at least one link
                if (externalLinks == null || ((ICollection)externalLinks).Count == 0)
                {
                    Console.WriteLine("No external links found in the workbook.");
                    return;
                }

                // Replace the first external link
                dynamic firstLink = externalLinks[0];
                string oldLink = firstLink.SourceFullName;

                try
                {
                    firstLink.ChangeExternalLink(newLinkPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to change external link: {ex.Message}");
                    return;
                }

                // Prepare backup path
                string backupPath = Path.Combine(
                    Path.GetDirectoryName(workbookPath) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(workbookPath) + "_backup.xlsx");

                // Save the workbook as a backup
                try
                {
                    wb.Save(backupPath, SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save backup workbook: {ex.Message}");
                    return;
                }

                Console.WriteLine("External link updated successfully.");
                Console.WriteLine($"Old link: {oldLink}");
                Console.WriteLine($"New link: {newLinkPath}");
                Console.WriteLine($"Backup saved to: {backupPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
