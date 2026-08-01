// Title: Protect a VBA Project in an XLSM Workbook and Verify with Aspose.Cells for .NET
// Description: Demonstrates how to add a VBA module to a new workbook, protect the VBA project with a password, save and reload the file, check the IsProtected flag, attempt to read the module's BinaryCodes (which fails when protected), and validate the password using Aspose.Cells APIs.
// Keywords: Aspose.Cells VBA protection | C# protect VBA project | read VBA module binary codes | validate VBA password | macro-enabled workbook | VbaProject.Protect | VbaProject.ValidatePassword | IsProtected property
// Common Searches: protect VBA project with Aspose.Cells .NET | check if VBA project is locked after saving | read VBA module code from protected XLSM | validate VBA password using Aspose.Cells
// Developer Intent: The developer wants to secure a VBA project with a password and confirm that the protection blocks access to the module source code.
// Use Cases: Secure macro source code before distributing an Excel file. | Programmatically detect whether a loaded workbook’s VBA project is protected. | Validate a VBA project password after opening a protected workbook. | Handle the exception thrown when attempting to read BinaryCodes of a locked module.
// AI Prompts: Generate C# code with Aspose.Cells to protect a VBA project and verify protection by checking IsProtected and attempting to read BinaryCodes. | Show how to catch and handle the exception when reading BinaryCodes of a protected VBA module in .NET. | Explain the steps to validate a VBA project password after loading a protected XLSM file using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a VBA module to a new workbook, protect the VBA project with a password, save and reload the file, check the IsProtected flag, attempt to read the module's BinaryCodes (which fails when protected), and validate the password using Aspose.Cells APIs.
    public class VbaProjectProtectionAndReadDemo
    {
        public static void Run()
        {
            try
            {
                // Step 1: Create a new workbook
                Workbook wb = new Workbook();

                // Step 2: Ensure a VBA project exists by saving as a macro‑enabled file and reloading
                string tempPath = "temp.xlsm";
                wb.Save(tempPath, SaveFormat.Xlsm);
                if (File.Exists(tempPath))
                {
                    wb = new Workbook(tempPath);
                    File.Delete(tempPath);
                }

                // Step 3: Add a VBA module to the project
                int moduleIndex = wb.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");

                // Step 4: Protect the VBA project and lock it for viewing
                string password = "SecretPwd";
                wb.VbaProject.Protect(true, password);

                // Step 5: Save the protected workbook
                string protectedPath = "ProtectedVbaProject.xlsm";
                wb.Save(protectedPath, SaveFormat.Xlsm);

                // Step 6: Load the protected workbook
                if (!File.Exists(protectedPath))
                {
                    Console.WriteLine($"File not found: {protectedPath}");
                    return;
                }

                Workbook loadedWb = new Workbook(protectedPath);
                VbaProject vbaProject = loadedWb.VbaProject;

                // Step 7: Output protection status
                Console.WriteLine("Is VBA Project Protected: " + vbaProject.IsProtected);
                // The IsLockedForViewing property is not available in this version of Aspose.Cells.
                // If needed, you can infer the lock state from other properties or methods.

                // Step 8: Attempt to read the binary codes of the first module
                try
                {
                    if (vbaProject.Modules.Count > 0)
                    {
                        byte[] binaryCodes = vbaProject.Modules[0].BinaryCodes;
                        if (binaryCodes != null && binaryCodes.Length > 0)
                        {
                            string codeContent = Encoding.UTF8.GetString(binaryCodes);
                            Console.WriteLine("Module code length: " + binaryCodes.Length);
                            Console.WriteLine("Module code preview:");
                            Console.WriteLine(codeContent.Substring(0, Math.Min(200, codeContent.Length)));
                        }
                        else
                        {
                            Console.WriteLine("Unable to read module code – likely due to protection.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("No VBA modules found in the project.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error reading module code: " + ex.Message);
                }

                // Step 9: Validate the password (should return true)
                bool isPasswordValid = vbaProject.ValidatePassword(password);
                Console.WriteLine("Password validation result: " + isPasswordValid);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectProtectionAndReadDemo.Run();
        }
    }
}
