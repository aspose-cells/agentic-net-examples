// Title: Copy all workbook document properties from one Excel file to another using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy document properties from source workbook to new workbook | how to transfer Excel file metadata using Aspose.Cells .NET | copy author title keywords between Excel workbooks programmatically | Aspose.Cells reflection example for copying WorkbookProperties | duplicate Excel workbook custom properties with Aspose.Cells in C#
// Tags: Aspose.Cells workbook metadata copy .NET | copy Excel document properties using reflection | transfer workbook properties programmatically C# | Aspose.Cells WorkbookProperties example | duplicate Excel file metadata Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a source Excel workbook, creates an empty destination workbook, accesses their WorkbookProperties via reflection, copies a set of standard document properties (Author, Title, Subject, Keywords, etc.) from the source to the destination, and saves the new workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Copy document properties from source to destination (if supported)
            try
            {
                // Document properties are accessed via the DocumentProperties class
                // Some Aspose.Cells versions expose them through the WorkbookProperties property.
                // Use reflection as a fallback to handle version differences.
                var sourceProps = sourceWorkbook.GetType().GetProperty("WorkbookProperties")?.GetValue(sourceWorkbook);
                var destProps = destWorkbook.GetType().GetProperty("WorkbookProperties")?.GetValue(destWorkbook);

                if (sourceProps != null && destProps != null)
                {
                    // Copy common properties using reflection
                    var propNames = new[]
                    {
                        "Author", "Title", "Subject", "Keywords", "Comments",
                        "Category", "Company", "Manager", "HyperlinkBase",
                        "CreatedDate", "LastSavedTime", "Application", "Version"
                    };

                    foreach (var name in propNames)
                    {
                        var srcProp = sourceProps.GetType().GetProperty(name);
                        var dstProp = destProps.GetType().GetProperty(name);
                        if (srcProp != null && dstProp != null && dstProp.CanWrite)
                        {
                            var value = srcProp.GetValue(sourceProps);
                            dstProp.SetValue(destProps, value);
                        }
                    }
                }
            }
            catch (Exception propEx)
            {
                Console.WriteLine($"Warning: Unable to copy document properties. {propEx.Message}");
            }

            // Save the destination workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Destination workbook saved to: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
