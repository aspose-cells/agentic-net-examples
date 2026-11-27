// Title: Copy rows that contain list‑type data validation while preserving dropdown rules with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy rows with list validation and keep dropdowns | Transfer rows with data‑validation lists between workbooks using Aspose.Cells | Verify that data validation rules remain after copying rows in Aspose.Cells | C# example for duplicating rows and preserving list‑type validation in Excel | How to duplicate validation objects for specific rows with Aspose.Cells
// Tags: copy rows with list validation Aspose.Cells | preserve dropdown validation C# | transfer data‑validation objects between worksheets | validate copied validation areas Aspose.Cells | use CopyOptions for validation duplication .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace TransferValidationRows
{
    // The sample loads a source workbook, detects every row that contains a list‑type data validation, copies those rows (including values, styles, and formulas) to a new workbook, recreates the corresponding validation objects only for the transferred rows using CopyOptions, verifies that each validation’s formula and cell area were correctly replicated, and saves the destination file.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // Input file – change the path as needed or pass it as a command line argument
            // -----------------------------------------------------------------
            string sourcePath = args.Length > 0 ? args[0] : "source.xlsx";

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source file \"{sourcePath}\" does not exist.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook sourceWb = new Workbook(sourcePath);
                Worksheet sourceWs = sourceWb.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Identify rows that contain List‑type data validation
                // -----------------------------------------------------------------
                HashSet<int> rowsWithListValidation = new HashSet<int>();
                ValidationCollection sourceValidations = sourceWs.Validations;

                foreach (Validation val in sourceValidations)
                {
                    if (val.Type != ValidationType.List)
                        continue; // we only care about list validations

                    foreach (CellArea area in val.Areas)
                    {
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            rowsWithListValidation.Add(r);
                        }
                    }
                }

                if (rowsWithListValidation.Count == 0)
                {
                    Console.WriteLine("No rows with List data validation were found in the source worksheet.");
                    return;
                }

                // Sort rows to keep original order
                List<int> rowsToCopy = rowsWithListValidation.OrderBy(r => r).ToList();

                // -----------------------------------------------------------------
                // 2. Create destination workbook and copy the identified rows
                // -----------------------------------------------------------------
                Workbook destWb = new Workbook(); // creates a default empty workbook
                Worksheet destWs = destWb.Worksheets[0];

                // Determine the maximum column used in the source sheet (simple heuristic)
                int maxColumn = sourceWs.Cells.MaxColumn;

                foreach (int srcRow in rowsToCopy)
                {
                    // Copy each cell in the row (value, style, formula, etc.)
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell srcCell = sourceWs.Cells[srcRow, col];
                        Cell destCell = destWs.Cells[srcRow, col];
                        destCell.Copy(srcCell);
                    }
                }

                // -----------------------------------------------------------------
                // 3. Copy the validation objects that belong to the copied rows
                // -----------------------------------------------------------------
                ValidationCollection destValidations = destWs.Validations;

                foreach (Validation srcVal in sourceValidations)
                {
                    if (srcVal.Type != ValidationType.List)
                        continue; // only list validations are relevant for this demo

                    // Determine whether this validation applies to any of the rows we copied
                    List<CellArea> intersectingAreas = new List<CellArea>();
                    foreach (CellArea area in srcVal.Areas)
                    {
                        // If any row of the area is in rowsToCopy, we keep the whole area
                        bool intersects = rowsToCopy.Any(r => r >= area.StartRow && r <= area.EndRow);
                        if (intersects)
                        {
                            intersectingAreas.Add(area);
                        }
                    }

                    if (intersectingAreas.Count == 0)
                        continue; // this validation does not affect the rows we transferred

                    // Add a new validation to the destination worksheet
                    int newIndex = destValidations.Add();
                    Validation destVal = destValidations[newIndex];

                    // Copy all properties from the source validation
                    destVal.Copy(srcVal, new CopyOptions());

                    // Clear any automatically added areas (Copy does not copy areas)
                    // and add only the intersecting ones
                    foreach (CellArea area in intersectingAreas)
                    {
                        destVal.AddArea(area);
                    }
                }

                // -----------------------------------------------------------------
                // 4. Verify that the validation rules are functional in the destination
                // -----------------------------------------------------------------
                Console.WriteLine("\nVerification of copied validations:");
                bool allOk = true;

                foreach (Validation srcVal in sourceValidations)
                {
                    if (srcVal.Type != ValidationType.List)
                        continue;

                    // Find the corresponding validation in the destination by comparing
                    // the formula and type (simple heuristic)
                    Validation destMatch = null;
                    foreach (Validation dVal in destValidations)
                    {
                        if (dVal.Type == srcVal.Type &&
                            dVal.Formula1 == srcVal.Formula1 &&
                            dVal.InCellDropDown == srcVal.InCellDropDown)
                        {
                            destMatch = dVal;
                            break;
                        }
                    }

                    if (destMatch == null)
                    {
                        Console.WriteLine($"[FAIL] Validation \"{srcVal.Formula1}\" was not copied.");
                        allOk = false;
                        continue;
                    }

                    // Verify that at least one area matches
                    bool areaMatch = false;
                    foreach (CellArea srcArea in srcVal.Areas)
                    {
                        foreach (CellArea dstArea in destMatch.Areas)
                        {
                            if (srcArea.StartRow == dstArea.StartRow &&
                                srcArea.StartColumn == dstArea.StartColumn &&
                                srcArea.EndRow == dstArea.EndRow &&
                                srcArea.EndColumn == dstArea.EndColumn)
                            {
                                areaMatch = true;
                                break;
                            }
                        }
                        if (areaMatch) break;
                    }

                    if (!areaMatch)
                    {
                        Console.WriteLine($"[FAIL] Validation \"{srcVal.Formula1}\" areas were not correctly transferred.");
                        allOk = false;
                    }
                    else
                    {
                        Console.WriteLine($"[PASS] Validation \"{srcVal.Formula1}\" copied successfully.");
                    }
                }

                // -----------------------------------------------------------------
                // 5. Save the destination workbook
                // -----------------------------------------------------------------
                string destPath = "destination.xlsx";
                destWb.Save(destPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nDestination workbook saved to \"{destPath}\".");

                if (allOk)
                {
                    Console.WriteLine("\nAll validations were transferred and verified successfully.");
                }
                else
                {
                    Console.WriteLine("\nSome validations failed verification – see messages above.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
