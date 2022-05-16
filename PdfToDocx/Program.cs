using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using Tesseract;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfToDocx
{
    class Program
    {
        public static List<CountyZipCodeInfo> CountyZipCodeInfoList;


        public static void Main(string[] args)
        {
            string json = File.ReadAllText(
                Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "ZipCodeMap.json"));
            CountyZipCodeInfoList = JsonSerializer.Deserialize<List<CountyZipCodeInfo>>(json);

            foreach (var path in args)
                PdfToEnvelopDocx(path);


            //PdfToEnvelopDocx("data\\成家美地整棟111H6000278REGA.pdf");

            //GetEnvelopInfoFromPdf();
            //Example01_WordTmplRendering();

            Console.WriteLine("Press any key to exit this program...");
            Console.Read();
        }


        private static string getZip(string fullAddress)
        {
            string zip = "";

            string[] seperators = new string[] { "市", "縣", "區", "鎮", "鄉" };

            var addressSegments = fullAddress.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
            string county = addressSegments[0];
            string district = addressSegments[1];

            foreach (var countyZipCodeInfo in CountyZipCodeInfoList.FindAll(i => i.Name.Contains(county)))
            {
                var finds = countyZipCodeInfo.Districts.FindAll(d => d.Name.Contains(district));
                if (finds.Count > 0)
                    zip = finds[0].Zip;
                else
                    Console.WriteLine($"{countyZipCodeInfo.Name} cannot find: {district}");
            }

            return zip;
        }


        public static void PdfToEnvelopDocx(string pdfFilePath)
        {
            string dir = Path.GetDirectoryName(pdfFilePath);
            string pdfFileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfFilePath);
            string templateFilePath = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "template.docx");
            string destFilePath = $"{dir}\\{pdfFileNameWithoutExt}-{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            string imgDir = Path.Combine(dir, $"{pdfFileNameWithoutExt}_image");

            List<PdfEnvelopItem> pdfEnvelopItemList = GetEnvelopInfoFromPdf(pdfFilePath);

            Console.WriteLine($"Generating tag-value map...");
            TesseractEngine ocr;
            ocr = new TesseractEngine("", "chi_tra", EngineMode.TesseractOnly);
            Dictionary<string, object> map = new Dictionary<string, object>();
            for (int i = 0; i < pdfEnvelopItemList.Count; i++)
            {
                Console.WriteLine($"Sequence: {pdfEnvelopItemList[i].Sequence}...");
                if (pdfEnvelopItemList[i].Address != null)
                {
                    map.Add($"address{i}", pdfEnvelopItemList[i].Address);
                    string zip = getZip(pdfEnvelopItemList[i].Address);
                    pdfEnvelopItemList[i].Zip = zip;
                    Console.WriteLine($"{zip}  {pdfEnvelopItemList[i].Address}");
                }
                else
                {
                    var imageData = new DocxHelper.ImageData(Path.Combine(imgDir, pdfEnvelopItemList[i].AddressImage.FileName));

                    Pix pix = Pix.LoadFromMemory(imageData.BinaryData);
                    Tesseract.Page tpage = ocr.Process(pix, PageSegMode.SingleBlock);
                    string addressTxt = tpage.GetText();
                    tpage.Dispose();
                    addressTxt = addressTxt.Replace(" ", "");
                    string zip = getZip(addressTxt);
                    pdfEnvelopItemList[i].Zip = zip;
                    Console.WriteLine($"{zip}  {addressTxt}");

                    imageData.Height = 0.6M;
                    imageData.Width = 15;
                    map.Add($"address{i}", imageData);
                }

                map.Add($"recipient{i}", pdfEnvelopItemList[i].Owner);
                map.Add($"zip{i}", pdfEnvelopItemList[i].Zip);
            }
            Console.WriteLine($"Generating tag-value map done.");
            
            Console.WriteLine($"Reading docx template and creating target temp docx...");
            using (var template = WordprocessingDocument.Open(templateFilePath, false))
            using (var dest = WordprocessingDocument.Create(destFilePath, WordprocessingDocumentType.Document))
            {
                Console.WriteLine($"building target temp docx content...");
                foreach (var part in template.Parts)
                    dest.AddPart(part.OpenXmlPart, part.RelationshipId);

                for (int i = 0; i < pdfEnvelopItemList.Count; i++)
                {
                    string contentStr = template.MainDocumentPart.Document.InnerXml;
                    contentStr = contentStr.Replace("[$recipient$]", $"[$recipient{i}$]")
                                            .Replace("[$zip$]", $"[$zip{i}$]")
                                            .Replace("[$address$]", $"[$address{i}$]");
                    if (i == 0)
                        dest.MainDocumentPart.Document.InnerXml = contentStr;
                    else
                        dest.MainDocumentPart.Document.InnerXml += contentStr;
                }
                Console.WriteLine($"building target temp docx content done.");
            }

            Console.WriteLine($"Generate target docx...");
            var docxBytes = DocxHelper.GenerateDocx(File.ReadAllBytes(destFilePath), map);
            File.WriteAllBytes(destFilePath, docxBytes);
            Console.WriteLine($"Writing target docx done.");

        }

        public static List<PdfEnvelopItem> GetEnvelopInfoFromPdf(string pdfPath)
        {
            string dir = Path.GetDirectoryName(pdfPath);
            string pdfFileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
            string imgDir = Path.Combine(dir, $"{pdfFileNameWithoutExt}_image");

            Console.WriteLine($"Read PDF: {pdfPath}");

            List<PdfEnvelopItem> pdfEnvelopItemList = new List<PdfEnvelopItem>();
            List<AddressImageItem> addressImageItemList = new List<AddressImageItem>();
            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                int pageNum = 1;

                Console.WriteLine($"Get text...");
                foreach (UglyToad.PdfPig.Content.Page page in document.GetPages())
                {
                    Console.Write($"\rpage: {pageNum}");
                    foreach (Word word in page.GetWords())
                    {
                        if (word.Text.StartsWith("（") && word.Text.Length == 15)
                            pdfEnvelopItemList.Add(new PdfEnvelopItem() { Page = pageNum, SequenceWord = word });
                        
                        if (word.Text.StartsWith("所有權人"))
                            pdfEnvelopItemList.Last().OwnerWord = word;

                        if (word.Text.StartsWith("統一編號"))
                            pdfEnvelopItemList.Last().IdWord = word;

                        if (word.Text.StartsWith("址："))
                            pdfEnvelopItemList.Last().AddressWord = word;

                        if (word.Text.StartsWith("權利範圍"))
                            pdfEnvelopItemList.Last().OtherWord = word;
                    }
                    pageNum++;
                }
                Console.WriteLine();
                Console.WriteLine($"Get text done.");

                pageNum = 1;
                Console.WriteLine($"Get image...");
                foreach (UglyToad.PdfPig.Content.Page page in document.GetPages()) 
                { 
                    int i = 1;
                    var pdfImages = page.GetImages();
                    pdfImages = pdfImages.OrderByDescending(i => i.Bounds.BottomLeft.Y);
                    foreach (var pdfImage in pdfImages)
                    {
                        Console.Write($"\rpage-index: {pageNum}-{i}");
                        byte[] png;
                        if (pdfImage.TryGetPng(out png))
                        {
                            using (MemoryStream ms = new MemoryStream(png))
                            using (Image image = Image.FromStream(ms, true, true))
                            {

                                Bitmap bitmap = ImageHelper.ImageTrim(new Bitmap(image));
                                if (bitmap.Width >= 230)
                                {
                                    //bitmap = ResizeImage(image, 1800, image.Height);
                                    bitmap = bitmap.Clone(new Rectangle(230, 0, bitmap.Width - 230, bitmap.Height), bitmap.PixelFormat);
                                }
                                if(bitmap.Height>40 && bitmap.Height < 50 && bitmap.Width / bitmap.Height > 10)
                                {
                                    string fileName = $"{pageNum}-{i}_{pdfImage.Bounds.BottomLeft.Y}.png";
                                    Directory.CreateDirectory(imgDir);
                                    bitmap.Save(Path.Combine(imgDir, $"{fileName}"), System.Drawing.Imaging.ImageFormat.Png);
                                    addressImageItemList.Add(new AddressImageItem() { Page = pageNum, FileName = fileName, AddressImage = pdfImage });
                                    i++;
                                }
                                else
                                {
                                    string fileName = $"{pageNum}-{i}.png";
                                    //Directory.CreateDirectory(imgDir);
                                    //bitmap.Save(Path.Combine(imgDir, $"______{fileName}"), System.Drawing.Imaging.ImageFormat.Png);
                                }
                            }
                        }
                    }
                    pageNum++;
                }
                Console.WriteLine();
                Console.WriteLine($"Get image done.");
            }

            Console.WriteLine($"Mapping envelop item information...");
            for (int i = 0; i < pdfEnvelopItemList.Count; i++)
            {
                var item = pdfEnvelopItemList[i];
                Console.Write($"\ritem sequence: {item.Sequence}");
                if (item.AddressWord != null)
                    continue;
                if (i == pdfEnvelopItemList.Count - 1)
                {
                    pdfEnvelopItemList[i].AddressImage = addressImageItemList.Last();
                }
                else
                {
                    double up = item.SequencePoint.Y;
                    double low = item.IdPoint.Y-30;
                    item.AddressImage = addressImageItemList.Find(
                        a => a.Page == item.Page && a.AddressImage.Bounds.BottomLeft.Y < up && a.AddressImage.Bounds.BottomLeft.Y > low);

                    if(item.AddressImage == null)
                    {
                        low = item.OtherWordPoint.Y;
                        item.AddressImage = addressImageItemList.Find(a => a.Page == item.Page + 1 && a.AddressImage.Bounds.BottomLeft.Y > low);
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine($"Mapping envelop item information done.");
            return pdfEnvelopItemList;

        }





    }
}
