// Title: Handle empty password exception when protecting a VBA project with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells protect VBA project with empty password without throwing exception | how to safely call workbook.VbaProject.Protect when password may be blank in .NET | catch exception from VbaProject.Protect for empty password using Aspose.Cells | skip VBA project protection in Aspose.Cells if password string is empty | error handling pattern for VBA project protection in Aspose.Cells C#
// Tags: VbaProject.Protect empty password handling | Aspose.Cells VBA project protection C# | exception handling workbook.VbaProject.Protect | skip VBA protection when password is null Aspose.Cells | load save workbook error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook, checking for an empty VBA password, and wrapping workbook.VbaProject.Protect in a try‑catch block to capture exceptions caused by a blank password. It also shows conditional skipping of protection and robust error handling for loading and saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook (lifecycle rule: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Password for VBA project protection (empty string will cause an exception)
        string password = ""; // Replace with a non‑empty password if protection is required

        if (string.IsNullOrEmpty(password))
        {
            Console.WriteLine("Password is empty; skipping VBA project protection.");
        }
        else
        {
            try
            {
                // Protect the VBA project (feature rule: VbaProject.Protect)
                // The first argument indicates whether the project is read‑only; false allows editing.
                workbook.VbaProject.Protect(false, password);
            }
            catch (Exception ex)
            {
                // Handle any exception that occurs during protection
                Console.WriteLine($"Failed to protect VBA project: {ex.Message}");
            }
        }

        try
        {
            // Save the workbook after the operation (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
