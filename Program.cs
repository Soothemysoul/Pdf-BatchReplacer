using Acrobat;
using Microsoft.Office.Interop.Word;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes;
using Spire.Pdf;
using Spire.Pdf.Exporting.XPS.Schema;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;
using Path = System.IO.Path;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace ReplaceTextInPdf
{
    class Program
    {

        public static AcroApp AcrobatApp { get; private set; }

        static void Main(string[] args)
        {
            try
            {
                AcrobatApp = new AcroApp();


                string executableFile = Assembly.GetExecutingAssembly().Location;
                DirectoryInfo directoryInfo = new DirectoryInfo(executableFile);

                if (directoryInfo.Parent.EnumerateDirectories().Count() == 0)
                {
                    Console.WriteLine("Папки не обнаружены!");
                    Console.ReadKey();
                    return;
                }

                foreach (var subdirectory in directoryInfo.Parent.EnumerateDirectories())
                {
                    try
                    {
                        int counter = 0;
                        foreach (string file in Directory.GetFiles(subdirectory.FullName, "*.pdf"))
                        {
                            counter++;
                        }

                        if (counter == 0) { continue; }

                        if (counter <= 2)
                        {
                            foreach (string file in Directory.GetFiles(subdirectory.FullName, "*.pdf"))
                            {
                                if (Path.GetFileName(file).Contains("лист") || Path.GetFileName(file).Contains("ИУЛ") || Path.GetFileName(file).EndsWith("УЛ.pdf")) { continue; }

                                Console.WriteLine($"Обработка файла '{Path.GetFileName(file)}'");
                                Console.WriteLine();
                                PdfReplacer(file);
                            }
                        }

                        if (counter > 2)
                        {
                            string file = Directory.GetFiles(subdirectory.FullName, "*ТЧ.pdf").First();

                            if (file != null)
                            {
                                Console.WriteLine($"Обработка файла '{Path.GetFileName(file)}'");
                                Console.WriteLine();
                                PdfReplacer(file);
                            }

                        }
                    }
                    catch { continue; }


                }


                AcrobatApp.Maximize(1);
                AcrobatApp.Exit();
                AcrobatApp.Exit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
            }

            Console.WriteLine("Готово!\nНажмите любую клавишу для выхода.");
            Console.ReadKey();

        }

        public static void PdfReplacer(string pdfFile)
        {

            string tmpPdf = pdfFile.Replace(Path.GetFileName(pdfFile), "tmp.pdf");
            string tmpWord = tmpPdf.Replace("pdf", "docx");

            CAcroAVDoc doc = new AcroAVDoc();


            if (doc.Open(pdfFile, null))
            {
                CAcroPDDoc pdfDoc = doc.GetPDDoc();
                pdfDoc.DeletePages(2, pdfDoc.GetNumPages() - 1);
                pdfDoc.Save(1, tmpPdf);
                Object jsoObj = pdfDoc.GetJSObject();

                Type jsType = jsoObj.GetType();
                //have to use acrobat javascript api because, acrobat
                object[] saveAsParam = { tmpWord, "com.adobe.acrobat.docx", "", false, false };
                jsType.InvokeMember("saveAs", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, jsoObj, saveAsParam, CultureInfo.InvariantCulture);

                pdfDoc.Close();


            }

            doc.Close(0);
            AcrobatApp.CloseAllDocs();

            SearchReplace(tmpWord);

            if (doc.Open(pdfFile, null))
            {
                CAcroAVDoc sourceDoc = new AcroAVDoc();
                CAcroPDDoc sourcePdfDoc;
                sourceDoc.Open(tmpPdf, null);
                sourcePdfDoc = sourceDoc.GetPDDoc();


                CAcroPDDoc pdfDoc = doc.GetPDDoc();
                pdfDoc.ReplacePages(0, sourcePdfDoc, 0, 2, 0);
                pdfDoc.Save(1, pdfFile);


                pdfDoc.Close();
                sourcePdfDoc.Close();
                sourceDoc.Close(0);

            }

            doc.Close(1);
            AcrobatApp.CloseAllDocs();

            File.Delete(tmpPdf);
            File.Delete(tmpWord);

            Console.WriteLine();

        }

        public static void SearchReplace(string filePath)
        {
            Application objWord = null;

            objWord = new Application();

            string[] searchStr = { "Новая", "через Неву\r", "пр.\r", @"через Неву", "Большого Смоленского", "ул. Коллонтай.", "пр. Обуховской", "Этап строительства №2", " пр." };
            string[] replaceStr = { "\"Новая", "через р.Неву ", "пр. ", "через р.Неву", "Б.Смоленского", "ул.Коллонтай.", "пр.Обуховской", "(1-й этап и 2-й этап)\" 2-й" + (char)160 + "этап", (char)160 + "пр." };

            // Активный документ
            Document objDoc = objWord.Documents.Open(filePath, ReadOnly: false);

            float shapeWidth = objWord.MillimetersToPoints(175.0f);
            float offsetLeft = (objDoc.PageSetup.PageWidth - shapeWidth) / 2;


            // Получение диапазона текста в документе
            Range objRange = objDoc.Range();

            Shapes shapes = objDoc.Shapes;

            int count = 0;

            foreach (string str in searchStr)
            {
                bool isFind = false;
                // Замена слова
                objRange.Find.ClearFormatting();
                objRange.Find.Replacement.ClearFormatting();
                objRange.Find.Text = str;
                objRange.Find.Replacement.Text = replaceStr[count];

                object replaceAll = WdReplace.wdReplaceAll;

                objRange.Find.Execute(Replace: replaceAll);

                isFind = objRange.Find.Found;

                foreach (Shape shape in shapes)
                {
                    try
                    {
                        string shapeStr = shape.TextFrame.TextRange.Text;

                        if (shapeStr.Contains(str))
                        {
                            shape.TextFrame.TextRange.Text = shapeStr.Replace(str, replaceStr[count]);
                            shape.Width = shapeWidth;
                            shape.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                            shape.Left = offsetLeft;

                            if (!isFind) { isFind = true; }
                        }

                    }
                    catch { }

                }


                // Обработка результата
                if (isFind)
                {
                    Console.WriteLine($"Вхождение '{str}' заменено на '{replaceStr[count]}'.");
                }
                else
                {
                    Console.WriteLine($"Вхождение '{str}' не найдено.");
                }

                count += 1;

            }



            objDoc.ExportAsFixedFormat(filePath.Replace("docx", "pdf"), WdExportFormat.wdExportFormatPDF);

            objDoc.Close(true);
            // Release resources
            objWord.Quit();

        }
    }
}