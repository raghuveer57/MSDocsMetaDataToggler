using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;

namespace OfficeDocumentUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0 || (args[0] != "toggle" && args[0] != "get"))
            {
                Console.WriteLine("Usage: OfficeDocumentUpdater <toggle|get>");
                return;
            }

            string operation = args[0];

            // Try to get active instances of Word, Excel, and PowerPoint
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.PowerPoint.Application pptApp = null;

            try
            {
                wordApp = (Microsoft.Office.Interop.Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch (COMException)
            {
                // Word is not open
            }

            try
            {
                excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                // Excel is not open
            }

            try
            {
                pptApp = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                // PowerPoint is not open
            }

            if (wordApp != null && IsDocumentActive(wordApp))
            {
                HandleWord(operation, wordApp);
            }
            else if (excelApp != null && IsWorkbookActive(excelApp))
            {
                HandleExcel(operation, excelApp);
            }
            else if (pptApp != null && IsPresentationActive(pptApp))
            {
                HandlePpt(operation, pptApp);
            }
            else
            {
                Console.WriteLine("No active Word, Excel, or PowerPoint application found.");
            }
        }

        static bool IsDocumentActive(Microsoft.Office.Interop.Word.Application wordApp)
        {
            try
            {
                var doc = wordApp.ActiveDocument;
                return doc != null;
            }
            catch
            {
                return false;
            }
        }

        static bool IsWorkbookActive(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            try
            {
                var wb = excelApp.ActiveWorkbook;
                return wb != null;
            }
            catch
            {
                return false;
            }
        }

        static bool IsPresentationActive(Microsoft.Office.Interop.PowerPoint.Application pptApp)
        {
            try
            {
                var ppt = pptApp.ActivePresentation;
                return ppt != null;
            }
            catch
            {
                return false;
            }
        }

        static void HandleWord(string operation, Microsoft.Office.Interop.Word.Application wordApp)
        {
            Document activeDocument = null;

            try
            {
                activeDocument = wordApp.ActiveDocument;
            }
            catch (Exception)
            {
                Console.WriteLine("No active document found in Word. Please open a Word document first.");
                return;
            }

            if (activeDocument != null)
            {
                try
                {
                    dynamic builtInProps = activeDocument.BuiltInDocumentProperties;
                    bool foundConfi = false;
                    foreach (dynamic prop in builtInProps)
                    {
                        if (prop.Name == "Comments" || prop.Name == "Company")
                        {
                            if (operation == "toggle")
                            {
                                prop.Value = prop.Value == "filesettomesuvag" ? "" : "filesettomesuvag";
                                Console.WriteLine($"{prop.Name} property toggled successfully.");
                            }
                            else if (operation == "get")
                            {
                                if (prop.Value == "filesettomesuvag")
                                {
                                    foundConfi = true;
                                }
                            }
                        }
                    }

                    if (operation == "get")
                    {
                        Console.WriteLine(foundConfi ? "1" : "0");
                    }
                    else if (operation == "toggle")
                    {
                        try
                        {
                            activeDocument.Saved = false;
                            Thread.Sleep(300);
                            activeDocument.Save();
                            Console.WriteLine("Document saved successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error saving document: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error updating properties: " + ex.Message);
                }
            }
        }

        static void HandleExcel(string operation, Microsoft.Office.Interop.Excel.Application excelApp)
        {
            Workbook activeWorkbook = null;

            try
            {
                activeWorkbook = excelApp.ActiveWorkbook;
            }
            catch (Exception)
            {
                Console.WriteLine("No active workbook found in Excel. Please open an Excel document first.");
                return;
            }

            if (activeWorkbook != null)
            {
                try
                {
                    dynamic builtInProps = activeWorkbook.BuiltinDocumentProperties;
                    bool foundConfi = false;
                    foreach (dynamic prop in builtInProps)
                    {
                        if (prop.Name == "Comments" || prop.Name == "Company")
                        {
                            if (operation == "toggle")
                            {
                                prop.Value = prop.Value == "filesettomesuvag" ? "" : "filesettomesuvag";
                                Console.WriteLine($"{prop.Name} property toggled successfully.");
                            }
                            else if (operation == "get")
                            {
                                if (prop.Value == "filesettomesuvag")
                                {
                                    foundConfi = true;
                                }
                            }
                        }
                    }

                    if (operation == "get")
                    {
                        Console.WriteLine(foundConfi ? "1" : "0");
                    }
                    else if (operation == "toggle")
                    {
                        try
                        {
                            activeWorkbook.Saved = false;
                            Thread.Sleep(300);
                            activeWorkbook.Save();
                            Console.WriteLine("Workbook saved successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error saving workbook: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error updating properties: " + ex.Message);
                }
            }
        }

        static void HandlePpt(string operation, Microsoft.Office.Interop.PowerPoint.Application pptApp)
        {
            Presentation activePresentation = null;

            try
            {
                activePresentation = pptApp.ActivePresentation;
            }
            catch (Exception)
            {
                Console.WriteLine("No active presentation found in PowerPoint. Please open a PowerPoint document first.");
                return;
            }

            if (activePresentation != null)
            {
                try
                {
                    dynamic builtInProps = activePresentation.BuiltInDocumentProperties;
                    bool foundConfi = false;
                    foreach (dynamic prop in builtInProps)
                    {
                        if (prop.Name == "Comments" || prop.Name == "Company")
                        {
                            if (operation == "toggle")
                            {
                                prop.Value = prop.Value == "filesettomesuvag" ? "" : "filesettomesuvag";
                                Console.WriteLine($"{prop.Name} property toggled successfully.");
                            }
                            else if (operation == "get")
                            {
                                if (prop.Value == "filesettomesuvag")
                                {
                                    foundConfi = true;
                                }
                            }
                        }
                    }

                    if (operation == "get")
                    {
                        Console.WriteLine(foundConfi ? "1" : "0");
                    }
                    else if (operation == "toggle")
                    {
                        try
                        {
                            activePresentation.Saved = Microsoft.Office.Core.MsoTriState.msoFalse;
                            Thread.Sleep(300);
                            activePresentation.Save();
                            Console.WriteLine("Presentation saved successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error saving presentation: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error updating properties: " + ex.Message);
                }
            }
        }
    }
}