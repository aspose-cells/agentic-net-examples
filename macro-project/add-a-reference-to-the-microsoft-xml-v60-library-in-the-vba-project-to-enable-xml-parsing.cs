// Title: Programmatically add a Microsoft XML 6.0 reference to a VBA project in an .xlsm workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add MSXML 6.0 reference to VBA project in an Excel macro file with Aspose.Cells C# | Aspose.Cells C# add VBA reference for Microsoft XML library | C# code to insert MSXML2 reference into .xlsm VBA project using Aspose.Cells | using reflection to create VbaReferenceInfo for MSXML in Aspose.Cells | update VBA project references in Excel workbook via Aspose.Cells .NET
// Tags: Aspose.Cells add VBA reference | MSXML6 reference in Excel macro workbook | C# VbaReferenceInfo reflection | update VBA project references .NET | save .xlsm with added VBA reference

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceExample
{
    // The example loads an .xlsm workbook, accesses its VbaProject, creates a VbaReferenceInfo object for the MSXML2 (Microsoft XML 6.0) library via reflection, sets the GUID and version, adds it to the VBA project's References collection, and saves the workbook preserving the new reference.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsm";
                const string outputPath = "output.xlsm";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project in the workbook
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("Warning: The workbook does not contain a VBA project.");
                }
                else
                {
                    try
                    {
                        // Attempt to create a VbaReferenceInfo instance via reflection
                        Type refInfoType = Type.GetType("Aspose.Cells.Vba.VbaReferenceInfo, Aspose.Cells.Vba");
                        if (refInfoType != null)
                        {
                            object refInfo = Activator.CreateInstance(refInfoType);
                            refInfoType.GetProperty("Name")?.SetValue(refInfo, "MSXML2");
                            refInfoType.GetProperty("Guid")?.SetValue(refInfo, new Guid("{88D969C5-F192-11D4-A65F-0040963251E5}"));
                            refInfoType.GetProperty("Major")?.SetValue(refInfo, 6);
                            refInfoType.GetProperty("Minor")?.SetValue(refInfo, 0);
                            refInfoType.GetProperty("IsRegistered")?.SetValue(refInfo, true);

                            // Add the reference to the VBA project's references collection
                            MethodInfo addMethod = vbaProject.References.GetType().GetMethod("Add");
                            addMethod?.Invoke(vbaProject.References, new[] { refInfo });

                            Console.WriteLine("VBA reference added successfully.");
                        }
                        else
                        {
                            Console.WriteLine("VbaReferenceInfo type not found; skipping reference addition.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add VBA reference: {ex.Message}");
                    }
                }

                // Save the workbook preserving the VBA project and any added references
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
