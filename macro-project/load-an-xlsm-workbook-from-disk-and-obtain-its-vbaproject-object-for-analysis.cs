// Title: Load an XLSM workbook and access its VBA project with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to verify a macro‑enabled Excel file exists, load it using Aspose.Cells, obtain the Workbook.VbaProject object, and display the project name, signing status, protection flag, and module count, with handling for workbooks that lack a VBA project.
// Keywords: Aspose.Cells | C# .NET | XLSM | VbaProject | macro-enabled workbook | read VBA project | VBA module count | signed macros | protected VBA | extract VBA code | Excel automation
// Common Searches: Aspose.Cells C# load XLSM workbook | Get VBA project information from macro-enabled Excel using Aspose.Cells | How to read VBA modules count with Aspose.Cells .NET | Check if VBA project is signed in an XLSM file with Aspose.Cells | Determine if VBA project is protected in Excel workbook C#
// Developer Intent: The developer wants to load a macro‑enabled Excel file and inspect its VBA project metadata.
// Use Cases: Validate that uploaded XLSM files contain signed macros before processing. | Enumerate VBA modules to extract or document macro code. | Detect protected VBA projects to decide whether to modify or remove macros. | Log VBA project metadata for compliance auditing. | Generate a summary of macro usage across a batch of workbooks.
// AI Prompts: Write C# code using Aspose.Cells to load an XLSM file and list all VBA module names. | Create a method that returns true if the VBA project in a loaded workbook is signed, otherwise false. | Provide an example that extracts the source code from each VBA module and saves each as a .bas file. | Generate a script that checks whether a VBA project is password‑protected and reports the result. | Show how to copy VBA modules from one workbook to another using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify a macro‑enabled Excel file exists, load it using Aspose.Cells, obtain the Workbook.VbaProject object, and display the project name, signing status, protection flag, and module count, with handling for workbooks that lack a VBA project.
    public class LoadVbaProjectDemo
    {
        public static void Run()
        {
            // Path to the macro-enabled workbook
            string filePath = "sample.xlsm";

            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Obtain the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Display VBA project information if present
                if (vbaProject != null)
                {
                    Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
                    Console.WriteLine($"Is Signed: {vbaProject.IsSigned}");
                    Console.WriteLine($"Is Protected: {vbaProject.IsProtected}");
                    Console.WriteLine($"Modules Count: {vbaProject.Modules.Count}");
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadVbaProjectDemo.Run();
        }
    }
}
