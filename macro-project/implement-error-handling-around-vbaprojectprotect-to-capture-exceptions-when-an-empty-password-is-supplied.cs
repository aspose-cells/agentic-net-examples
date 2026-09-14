// Title: Implement try‑catch around VbaProject.Protect to handle empty password errors in Aspose.Cells for .NET
// AI Prompts: Generate C# code that wraps workbook.VbaProject.Protect in a try‑catch block and logs the exception when an empty password is supplied. | Create a helper method that checks the password string before calling VbaProject.Protect and throws a custom error if the password is blank, using Aspose.Cells.
// Common Searches: Aspose.Cells C# protect VBA project with empty password throws exception | how to catch error from VbaProject.Protect when password is blank in .NET | C# try catch for Aspose.Cells VBA project protection failure | handling Aspose.Cells VbaProject.Protect invalid password exception
// Tags: Aspose.Cells VBA project protection exception handling | C# try-catch Aspose.Cells VbaProject.Protect | empty password error Aspose.Cells VBA | protect VBA project without password .NET | error handling workbook VbaProject.Protect

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook or creates a new one, then attempts to protect its VBA project with an empty password. The Protect call is enclosed in a try‑catch block to capture and display any exceptions, ensuring robust error handling before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load existing workbook or create a new one if the file is missing
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
                Console.WriteLine($"Input file \"{inputPath}\" not found. A new workbook has been created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Protect the VBA project (read‑only flag set to false, empty password)
        try
        {
            workbook.VbaProject.Protect(false, "");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error protecting VBA project: {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
