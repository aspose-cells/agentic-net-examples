// Title: Check if an Excel workbook's structure or window is protected using Aspose.Cells for .NET (C#)
// Description: A concise C# example that loads an .xlsx file with Aspose.Cells, reads the WorkbookSettings.IsProtected flag, and prints the protection status (true/false) to the console. Demonstrates how to detect workbook structure or window protection without altering the file.
// Keywords: Aspose.Cells | C# | .NET | WorkbookSettings | IsProtected | workbook protection | structure protection | window protection | detect Excel protection | read protection status | encryption and protection
// Common Searches: Aspose.Cells check workbook protection | C# get workbook structure protected flag | How to read IsProtected property Aspose.Cells | Determine if Excel file is protected using Aspose.Cells | WorkbookSettings.IsProtected example
// Developer Intent: Identify whether the loaded workbook has structure or window protection enabled.
// Use Cases: Skip automated modifications when a workbook is protected | Log protection status of each file for compliance auditing | Conditionally unprotect a workbook before editing | Batch audit of Excel files to verify protection settings | Integrate protection checks into CI/CD pipelines
// AI Prompts: Generate C# code with Aspose.Cells that returns a boolean indicating workbook protection and writes a detailed log entry. | Create a reusable method that accepts a file path, checks WorkbookSettings.IsProtected, and throws a custom exception if protection is detected. | Write a PowerShell script that invokes a .NET assembly to report the protection status of multiple Excel files. | Explain how to differentiate between structure and window protection using Aspose.Cells properties.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionCheck
{
    // A concise C# example that loads an .xlsx file with Aspose.Cells, reads the WorkbookSettings.IsProtected flag, and prints the protection status (true/false) to the console. Demonstrates how to detect workbook structure or window protection without altering the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined.
            // Replace with the actual file path as needed.
            string filePath = "input.xlsx";

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Access the workbook settings.
            WorkbookSettings settings = workbook.Settings;

            // Determine if the workbook's structure or window is protected.
            bool isProtected = settings.IsProtected;

            // Output the protection status to the console.
            Console.WriteLine("Workbook structure/window protected: " + isProtected);
        }
    }
}
