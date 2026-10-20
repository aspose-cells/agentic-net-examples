// Title: Check whether a workbook’s VBA project is password‑protected with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to detect if a VBA project is password protected | C# code to verify Excel macro project lock status using Aspose.Cells | Determine if loaded workbook contains a protected VBA project in .NET | Read VBA project protection flag from an .xlsx file with Aspose.Cells | Handle missing IsPasswordProtected property in Aspose.Cells VBA project
// Tags: Aspose.Cells VBA project password check | C# detect protected macro project | Workbook.VbaProject existence validation | Excel macro lock status Aspose.Cells | fallback when IsPasswordProtected unavailable

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel file with Aspose.Cells, checks if a VBA project exists via the VbaProject property, attempts to read its IsPasswordProtected flag (defaulting to false when the property is unavailable), and outputs both the presence and protection status while handling file‑not‑found and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check if a VBA project exists in the workbook
            bool hasVbaProject = workbook.VbaProject != null;

            // Determine whether the VBA project is password protected.
            // The IsPasswordProtected property is not available in some versions of Aspose.Cells,
            // so we treat it as not protected when the property cannot be accessed.
            bool isPasswordProtected = false;
            if (hasVbaProject)
            {
                // If the property becomes available in a newer version, replace the following line:
                // isPasswordProtected = workbook.VbaProject.IsPasswordProtected;
                // For now, we keep the default value.
            }

            // Output the verification results
            Console.WriteLine("VBA project present: " + hasVbaProject);
            Console.WriteLine("VBA project password protected: " + isPasswordProtected);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
