// Title: Programmatically obtain the spilled range address of a dynamic array formula in cell C3 with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to get the spilled range address of the dynamic array located in cell C3, handling version differences with dynamic binding. | Show how to invoke GetSpilledRange and GetAddress on a worksheet cell via dynamic objects and catch RuntimeBinderException if the methods are unavailable. | Create a console app that loads an Excel file, checks whether C3 has a spilled range, and prints the address or a fallback message.
// Common Searches: Aspose.Cells C# retrieve spilled range address for dynamic array formula in cell C3 | How to use GetSpilledRange with Aspose.Cells when the method may not exist | Dynamic binding for Excel cell methods in Aspose.Cells .NET example | Check if cell C3 has a spilled range using Aspose.Cells and get its address | RuntimeBinderException handling for GetSpilledRange in Aspose.Cells
// Tags: Aspose.Cells GetSpilledRange dynamic binding | retrieve spilled range address C# | dynamic array formula Excel Aspose.Cells | handle RuntimeBinderException Aspose.Cells | load workbook access cell C3 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

// Loads an Excel workbook, accesses cell C3, uses dynamic binding to call GetSpilledRange and GetAddress, prints the spilled range address, and gracefully handles RuntimeBinderException when the API is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Use dynamic to call GetSpilledRange (available in newer versions) without compile‑time binding
            dynamic cell = worksheet.Cells["C3"];
            dynamic spilledRange = null;

            try
            {
                spilledRange = cell.GetSpilledRange();
            }
            catch (RuntimeBinderException)
            {
                // Method not available in the current Aspose.Cells version
                Console.WriteLine("GetSpilledRange method is not supported by the loaded Aspose.Cells version.");
            }

            if (spilledRange != null)
            {
                // Use dynamic again for GetAddress (may also be version‑specific)
                string spilledAddress = string.Empty;
                try
                {
                    spilledAddress = spilledRange.GetAddress();
                }
                catch (RuntimeBinderException)
                {
                    Console.WriteLine("GetAddress method is not supported by the loaded Aspose.Cells version.");
                }

                if (!string.IsNullOrEmpty(spilledAddress))
                {
                    Console.WriteLine($"Spilled range address: {spilledAddress}");
                }
                else
                {
                    Console.WriteLine("Unable to retrieve spilled range address.");
                }
            }
            else
            {
                Console.WriteLine("Cell C3 does not have a spilled range or the operation is unsupported.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
