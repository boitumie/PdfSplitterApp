using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using PdfSharpCore.Drawing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using UglyToad.PdfPig;

class Program
{
    [STAThread]
    static void Main()
    {
        Console.WriteLine("=== Split PDF and Save Only Labels with Matching Text, Then Print ===");

        string inputPath = ShowFileDialog();
        if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
        {
            Console.WriteLine("No valid PDF selected. Exiting...");
            return;
        }

        string outputDir = Path.Combine(Path.GetDirectoryName(inputPath)!, "Filtered_TextBeforeNewline_Labels");
        Directory.CreateDirectory(outputDir);

        
        List<string> pdfFilesToPrint = SplitAndFilterPdf(inputPath, outputDir);

        if (pdfFilesToPrint.Count == 0)
        {
            Console.WriteLine("No labels with matching text found to print.");
            return;
        }

       
        string selectedPrinter = SelectPrinter();
        if (string.IsNullOrEmpty(selectedPrinter))
        {
            Console.WriteLine("No printer selected. Exiting...");
            return;
        }

        
        foreach (var pdfFile in pdfFilesToPrint)
        {
            Console.WriteLine($"Printing {Path.GetFileName(pdfFile)} to printer: {selectedPrinter}");
            PrintPdf(pdfFile, selectedPrinter);
        }

        Console.WriteLine("All selected PDFs sent to printer.");
    }

    static List<string> SplitAndFilterPdf(string inputPath, string outputDir)
    {
        var savedPdfPaths = new List<string>();
        var seenLabels = new HashSet<string>();

        using var inputDoc = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);

        double labelHeight = 209.1; 
        int partCounter = 1;

        for (int pageIndex = 0; pageIndex < inputDoc.PageCount; pageIndex++)
        {
            var srcPage = inputDoc.Pages[pageIndex];
            double pageHeight = srcPage.Height;
            double pageWidth = srcPage.Width;

            for (double y = pageHeight; y > 0; y -= labelHeight)
            {
                double sliceHeight = Math.Min(labelHeight, y);
                double cropY = y - sliceHeight;

                using var outDoc = new PdfSharpCore.Pdf.PdfDocument();

           
                var newPage = outDoc.AddPage(srcPage); []

           
                var cropRect = new PdfSharpCore.Pdf.PdfRectangle(new XRect(0, cropY, pageWidth, sliceHeight));
                newPage.MediaBox = cropRect;
                newPage.CropBox = cropRect;
                newPage.TrimBox = cropRect;
               

                string tempPath = Path.Combine(outputDir, $"TEMP_Part_{partCounter}.pdf");
                
                outDoc.Save(tempPath);

                if (HasReadableText(tempPath, seenLabels))
                {
                    string finalPath = Path.Combine(outputDir, $"TextLabel_Part_{partCounter}.pdf");

                    if (File.Exists(finalPath))
                    {


                        File.Delete(finalPath);
                        File.Move(tempPath, finalPath);
                   
                        savedPdfPaths.Add(finalPath);
                        Console.WriteLine($"Saved: {finalPath}");
                    }
                    }
                    else
                {
                    File.Delete(tempPath);
                }

                partCounter++;
            }
        }

        return savedPdfPaths;
    }

    static bool HasReadableText(string pdfPath, HashSet<string> seenLabels)
    {
        using var doc = UglyToad.PdfPig.PdfDocument.Open(pdfPath);

        foreach (var page in doc.GetPages())
        {
            string text = page.Text?.Trim();
            if (string.IsNullOrEmpty(text)) continue;

            var matches = Regex.Matches(text, @"\*ST\d{3}R\dT\w{4} \d{4}\*");

            foreach (Match match in matches)
            {
                string label = match.Value;
                if (!seenLabels.Contains(label))
                {
                    seenLabels.Add(label);
                    return true;
                }
            }
        }

        return false;
    }

    static string SelectPrinter()
    {
        using var pd = new PrintDialog
        {
            AllowSomePages = false,
            AllowPrintToFile = false,
            UseEXDialog = true,
            
        };

        if (pd.ShowDialog() == DialogResult.OK)
        {

            return pd.PrinterSettings.PrinterName;
        }
        return null;
    }

    static void PrintPdf(string pdfFilePath, string printerName)
    {
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo()
            {
                FileName = pdfFilePath,
                Verb = "printto",
                Arguments = $"\"{printerName}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = true,
                
                
            };
            var proc = Process.Start(psi);

      
            proc.WaitForExit(10000);
            proc.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error printing file {pdfFilePath}: {ex.Message}");
        }
    }

    static string ShowFileDialog()
    {
        using var ofd = new OpenFileDialog
        {
            Filter = "PDF files (*.pdf)|*.pdf",
            Title = "Select PDF Document"
        };
        return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : string.Empty;
    }
}
