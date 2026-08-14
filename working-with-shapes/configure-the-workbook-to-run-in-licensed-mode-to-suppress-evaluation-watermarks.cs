// Title: C# – Apply Aspose.Cells .NET License to Suppress Evaluation Watermarks and Save Workbook
// Description: Demonstrates how to load an Aspose.Cells .NET license file, activate it with License.SetLicense, verify the activation via Workbook.IsLicensed, create a simple workbook, and save it as an Excel file without evaluation watermarks. Includes error handling for missing license files and I/O failures.
// Keywords: Aspose.Cells | .NET | C# | license | SetLicense | IsLicensed | remove watermark | evaluation watermark | save workbook | Excel file
// Common Searches: how to apply Aspose.Cells license in C# | remove Aspose.Cells evaluation watermark | Aspose.Cells SetLicense example | check if Aspose.Cells license is active | save licensed workbook with Aspose.Cells
// Developer Intent: Activate an Aspose.Cells .NET license to generate a watermark‑free workbook and persist it to disk.
// Use Cases: Load a .lic file from a known path and call License.SetLicense to enable full functionality. | Confirm licensing status with Workbook.IsLicensed before performing any spreadsheet operations. | Create and populate a workbook even when the license file is absent, logging the condition without crashing. | Save the workbook to a specified location while handling file‑system exceptions.
// AI Prompts: Generate C# code that reads an Aspose.Cells license from an embedded resource, applies it, and checks the license status. | Show a robust pattern for applying an Aspose.Cells .NET license with fallback logic when the license file cannot be found.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an Aspose.Cells .NET license file, activate it with License.SetLicense, verify the activation via Workbook.IsLicensed, create a simple workbook, and save it as an Excel file without evaluation watermarks. Includes error handling for missing license files and I/O failures.
public class LicensedWorkbookDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Apply the Aspose.Cells license to suppress evaluation watermarks
        try
        {
            string licensePath = "Aspose.Cells.NET.lic";
            if (File.Exists(licensePath))
            {
                License license = new License();
                license.SetLicense(licensePath);
                Console.WriteLine("License applied successfully.");
            }
            else
            {
                Console.WriteLine("License file not found. Continuing without a license.");
            }
        }
        catch (Exception licEx)
        {
            Console.WriteLine($"License error: {licEx.Message}");
        }

        // Verify that the license has been applied
        Console.WriteLine($"IsLicensed: {new Workbook().IsLicensed}");

        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets[0].Cells[0, 0].PutValue("Licensed Workbook");

        // Save the workbook to disk
        try
        {
            string outputPath = "LicensedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
        }
    }
}
