// Title: Reset workbook globalization to English (en-US) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to change workbook culture to en-US | reset Aspose.Cells workbook globalization settings to default English | set workbook CultureInfo to English in Aspose.Cells .NET | programmatically enforce en-US localization for an Excel workbook with Aspose.Cells | Aspose.Cells reset workbook localization for specific worksheets C#
// Tags: Aspose.Cells workbook CultureInfo en-US | C# enforce English localization Aspose.Cells | Aspose.Cells default globalization settings | set workbook culture Aspose.Cells .NET | globalization configuration workbook C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Provides a C# helper method that validates a Workbook, accesses its Settings, and assigns CultureInfo("en-US") to enforce English (United States) globalization for the entire workbook, with sample code for loading, resetting, and saving an Excel file using Aspose.Cells.
public static class GlobalizationHelper
{
    // Resets globalization settings to default English (en-US) for the entire workbook.
    public static void ResetToEnglish(Workbook workbook)
    {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        // Access the workbook's settings.
        WorkbookSettings settings = workbook.Settings;

        // Set the workbook culture to English (United States).
        settings.CultureInfo = new CultureInfo("en-US");

        // Additional system‑dependent format settings are not exposed in the current API.
        // If needed, they can be configured via other workbook options.
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Load the workbook (lifecycle rule).
            Workbook workbook = LoadWorkbook("input.xlsx");

            // Reset globalization settings for the workbook.
            GlobalizationHelper.ResetToEnglish(workbook);

            // Save the modified workbook (lifecycle rule).
            SaveWorkbook(workbook, "output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Loads a workbook from the specified path, ensuring the file exists.
    private static Workbook LoadWorkbook(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"The file '{path}' was not found.", path);

        try
        {
            return new Workbook(path);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to load workbook from '{path}'.", ex);
        }
    }

    // Saves a workbook to the specified path, handling any I/O errors.
    private static void SaveWorkbook(Workbook workbook, string path)
    {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        try
        {
            workbook.Save(path);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save workbook to '{path}'.", ex);
        }
    }
}
