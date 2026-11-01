// Title: How to validate that required smart markers exist in an Excel template using Aspose.Cells for .NET
// Common Searches: C# code to ensure all required smart markers are present in an Aspose.Cells Excel template before merging | Aspose.Cells validate missing smart markers in a workbook programmatically | check for &= smart markers in an Excel file using Aspose.Cells .NET | detect absent smart markers in an Excel template with Aspose.Cells C# | pre‑merge validation of smart markers in an Aspose.Cells template
// Tags: smart marker validation Aspose.Cells | required smart markers detection .NET | excel template marker presence check | Aspose.Cells workbook cell scanning | pre‑merge smart marker verification

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The SmartMarkerValidator class loads an Excel workbook, scans every cell for markers prefixed with "&=", collects their names, and compares them against a supplied list of required markers. It throws an InvalidOperationException listing any missing markers and writes a success message to the console when all required markers are found.
public class SmartMarkerValidator
{
    /// <param name="templatePath">Full path to the Excel template file.</param>
    /// <param name="requiredMarkers">List of smart marker names that must be present (e.g., "Customers.Name").</param>
    public static void Validate(string templatePath, List<string> requiredMarkers)
    {
        // Ensure the template file exists to avoid FileNotFoundException.
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        try
        {
            // Load the workbook (template) from the specified path.
            Workbook workbook = new Workbook(templatePath);

            // Collect smart markers present in the workbook by scanning cell values.
            HashSet<string> foundMarkers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        string cellText = cells[row, col].StringValue;
                        if (string.IsNullOrEmpty(cellText))
                            continue;

                        // Smart markers start with "&=" (e.g., "&=Customers.Name").
                        if (cellText.StartsWith("&="))
                        {
                            string marker = cellText.Substring(2); // Remove the leading "&="
                            // In case the marker contains additional formatting, take the first token.
                            int spaceIndex = marker.IndexOf(' ');
                            if (spaceIndex > 0)
                                marker = marker.Substring(0, spaceIndex);

                            foundMarkers.Add(marker);
                        }
                    }
                }
            }

            // Find any required markers that are missing.
            List<string> missingMarkers = requiredMarkers
                .Where(m => !foundMarkers.Contains(m))
                .ToList();

            // If there are missing markers, raise an informative exception.
            if (missingMarkers.Count > 0)
            {
                string missing = string.Join(", ", missingMarkers);
                throw new InvalidOperationException(
                    $"The following required smart markers are missing in the template '{templatePath}': {missing}");
            }

            // All required markers are present; validation succeeded.
            Console.WriteLine("Smart marker validation passed. All required markers are present.");
        }
        catch (Exception ex)
        {
            // Log the error and rethrow for upstream handling if needed.
            Console.WriteLine($"Validation failed: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Example usage:
        var required = new List<string> { "Customers.Name", "Customers.Address", "Orders.Total" };
        string templatePath = @"C:\Templates\ReportTemplate.xlsx";

        try
        {
            SmartMarkerValidator.Validate(templatePath, required);
        }
        catch (Exception ex)
        {
            // Handle or log the exception as appropriate.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
