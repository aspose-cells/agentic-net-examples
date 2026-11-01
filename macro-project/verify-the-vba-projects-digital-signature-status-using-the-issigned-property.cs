// Title: Check if an Excel workbook's VBA project is digitally signed using Aspose.Cells for .NET (IsSigned property)
// Common Searches: C# Aspose.Cells how to determine if a VBA project is signed in an .xlsx file | using Aspose.Cells .NET to read VBA digital signature status from an Excel workbook | sample code for checking VBA project IsSigned property with Aspose.Cells | detect signed macro in Excel using Aspose.Cells LoadOptions C# | verify VBA project digital signature presence in workbook with Aspose.Cells
// Tags: Aspose.Cells VBA digital signature check | C# IsSigned property Excel macro | load workbook inspect VBA signature Aspose | detect signed VBA project .NET | handle missing VBA project Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, checks for the presence of a VBA project, and uses the VbaProject.IsSigned property to report whether the VBA project is digitally signed, while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied)
            var loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Check if the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // Verify if the VBA project has a digital signature
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protected files)
            Console.WriteLine("Aspose.Cells error: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
