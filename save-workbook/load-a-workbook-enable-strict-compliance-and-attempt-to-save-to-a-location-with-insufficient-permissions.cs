// Title: Load an Excel workbook and handle UnauthorizedAccessException when saving to a protected folder with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook to C:\Windows\System32 and catch permission denied error | How to handle UnauthorizedAccessException while saving an Excel file using Aspose.Cells in .NET | Attempt to save an Aspose.Cells workbook to a directory without write access and retrieve error details | What exception is thrown by Aspose.Cells when saving to a restricted folder in C# | Load workbook with Aspose.Cells and emulate strict compliance behavior in .NET
// Tags: Aspose.Cells workbook load C# | Aspose.Cells save to protected directory | UnauthorizedAccessException handling Aspose.Cells | SaveFormat.Xlsx permission error | strict compliance mode Aspose.Cells alternative

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, tries to save it to a system‑protected folder, and catches UnauthorizedAccessException to report insufficient permissions, while noting that the StrictMode property is not available in the current Aspose.Cells version.
class Program
{
    static void Main()
    {
        // Path to the source workbook (ensure this file exists)
        string sourcePath = @"C:\Temp\input.xlsx";

        // Verify that the source file exists before attempting to load it
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(sourcePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // The StrictMode property is not available in the current Aspose.Cells version.
        // If needed, configure other relevant settings here.

        // Destination path with likely insufficient permissions
        string restrictedPath = @"C:\Windows\System32\restricted_output.xlsx";

        try
        {
            // Attempt to save the workbook to the restricted location
            workbook.Save(restrictedPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (UnauthorizedAccessException ex)
        {
            // Expected exception due to insufficient permissions
            Console.WriteLine("Failed to save workbook: insufficient permissions.");
            Console.WriteLine("Exception message: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
