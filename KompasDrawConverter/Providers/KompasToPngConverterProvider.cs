using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace KompasDrawConverter.Providers
{
    public class KompasToPngConverterProvider : KompasFileConverter
    {
        private readonly string _outputPngDirectory;
        private readonly string _outputPdfDirectory;

        public KompasToPngConverterProvider(Options options) : base(options)
        {
            _outputPngDirectory = Path.Combine(options.OutputDirectory, "png");
            _outputPdfDirectory = Path.Combine(options.OutputDirectory, "pdf");
        }

        protected override void CreateDirectories()
        {
            if (!Directory.Exists(_outputPngDirectory))
            {
                Directory.CreateDirectory(_outputPngDirectory);
            }
            
            if (!Directory.Exists(_outputPdfDirectory))
            {
                Directory.CreateDirectory(_outputPdfDirectory);
            }
        }

        protected override string[] GetAllowedFilesExtensions()
        {
            return new[] {"cdw", "spw", "m3d" };
        }

        protected override void ConvertFile(KompasInstance kompasInstance, string filePath)
        {
            var imagesFolder = kompasInstance.ConvertToRaster(filePath, _outputPngDirectory);

            ConvertToPdf(imagesFolder);
        }

        private void ConvertToPdf(string imagesFolder)
        {
            var pdfNameWithoutExt = Path.GetFileName(imagesFolder);
            var outputPdfName = $"{pdfNameWithoutExt}.pdf";
            var pdfPath = Path.Combine(_outputPdfDirectory, outputPdfName);

            var imagesDirectoryInfo = new DirectoryInfo(imagesFolder);
            var images = imagesDirectoryInfo.GetFiles();

            if (images.Length == 1)
            {
                CreatePdfFile(images[0].FullName, pdfPath);
            }
            else
            {
                var tempPdfFolderPath = Path.Combine(_outputPdfDirectory, "temp", pdfNameWithoutExt);
                Directory.CreateDirectory(tempPdfFolderPath);
                
                CreatePdfFiles(images, tempPdfFolderPath, pdfPath);
            }
        }

        private static void CreatePdfFiles(IEnumerable<FileInfo> images, string pdfTempFolderPath, string summaryPdfPath)
        {
            foreach (var imageInfo in images)
            {
                var imagePath = imageInfo.FullName;
                var imagePathWithoutExt = Path.GetFileNameWithoutExtension(imagePath);

                var pdfName = $"{imagePathWithoutExt}.pdf";
                var pdfPath = Path.Combine(pdfTempFolderPath, pdfName);

                CreatePdfFile(imagePath, pdfPath);
            }
            
            var pdfTempDirectory = new DirectoryInfo(pdfTempFolderPath);
            var tempPdfFiles = pdfTempDirectory.GetFiles().Select(x => x.FullName);
            MergePdfs(tempPdfFiles, summaryPdfPath);
        }

        private static void CreatePdfFile(string imagePath, string pdfPath)
        {
            iTextSharp.text.Rectangle pageSize;

            var width = 0;
            var height = 0;
            
            using (var srcImage = new Bitmap(imagePath))
            {
                width = srcImage.Width;
                height = srcImage.Height;
                
                pageSize = new iTextSharp.text.Rectangle(0, 0, width * 0.48f, height * 0.48f);
            }
            using (var ms = new MemoryStream())
            {
                var document = new Document(pageSize);
                PdfWriter.GetInstance(document, ms);
                document.Open();
                var image = iTextSharp.text.Image.GetInstance(imagePath);
                image.ScaleToFit(width * 0.48f, height * 0.48f);
                image.SetAbsolutePosition(0, 0);
                
                document.Add(image);
                document.Close();

                File.WriteAllBytes(pdfPath, ms.ToArray());
            }
        }

        private static void MergePdfs(IEnumerable<string> fileNames, string targetFileName)
        {
            using (var stream = new FileStream(targetFileName, FileMode.Create))
            {
                var document = new Document();
                var pdf = new PdfCopy(document, stream);
                PdfReader reader = null;

                try
                {
                    document.Open();
                    foreach (var file in fileNames)
                    {
                        reader = new PdfReader(file);
                        pdf.AddDocument(reader);
                        reader.Close();
                    }
                }
                catch (Exception)
                {
                    reader?.Close();
                }
                finally
                {
                    document.Close();
                }
            }
        }
    }
}

