// Title: Check if an Excel workbook requires an opening password or a modify password using Aspose.Cells for .NET
// Common Searches: how to verify if an Excel file is protected with an open password using Aspose.Cells C# | Aspose.Cells C# detect modify (write) password on a workbook | determine workbook password protection type Aspose.Cells .NET example | read write password property of WorkbookSettings with Aspose.Cells | C# load Excel file with Aspose.Cells and check password requirements
// Tags: read workbook open password Aspose.Cells | detect modify password WorkbookSettings Aspose.Cells | reflection access WritePassword property Aspose.Cells | handle missing Excel file Aspose.Cells | log password protection status C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, checks Workbook.Settings.Password for an opening password, uses reflection to read the optional WritePassword property for a modify password, and writes both statuses to the console while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        // Path to the workbook to be examined
        string filePath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied)
            Workbook workbook = new Workbook(filePath);

            // Determine if an opening password is set
            bool hasOpenPassword = !string.IsNullOrEmpty(workbook.Settings.Password);

            // Determine if a modify (write) password is set
            // Note: In some older Aspose.Cells versions the modify password property may not be available.
            // If unavailable, this check will default to false.
            bool hasModifyPassword = false;
            try
            {
                // Attempt to read the WritePassword property if it exists
                var writePasswordProp = typeof(WorkbookSettings).GetProperty("WritePassword");
                if (writePasswordProp != null)
                {
                    string writePassword = writePasswordProp.GetValue(workbook.Settings) as string;
                    hasModifyPassword = !string.IsNullOrEmpty(writePassword);
                }
            }
            catch
            {
                // If reflection fails, keep hasModifyPassword as false
            }

            // Log the results
            Console.WriteLine($"Open password required: {hasOpenPassword}");
            Console.WriteLine($"Modify password required: {hasModifyPassword}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
