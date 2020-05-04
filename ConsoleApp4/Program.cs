using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace ConsoleApp4
{ 
class Program
    {

        public static XDocument ExtractStylesPart( //Извлекаем стили оформления из самого дока
         string fileName,
         bool getStylesWithEffectsPart = true)
        {
            XDocument styles = null;
        
            // Открываем док и работаем с ним
            using (var document =
                WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
        
                //Проверка на стиль оформления дока, старый(до 2007) или новый
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                    stylesPart = docPart.StyleDefinitionsPart;
                else
                    stylesPart = docPart.StylesWithEffectsPart;
        
                if (stylesPart != null)
                {
                    using (var reader = XmlReader.Create(
                      stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        string fileNamejSon = @"C:\Users\DIMA\Downloads\Sample2.json"; //Куда сохраняется jSon файл
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(string));
                        MemoryStream stream = new MemoryStream();
                        JavaScriptSerializer serializer1 = new JavaScriptSerializer();
                        string json = serializer1.Serialize(styles.ToString());
                                           
                        File.WriteAllText(fileNamejSon, json);
                    }
                }
            }
            return styles;
        }


        //Извлекаем текст из дока
        public static string TextFromWord(string filename)
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            StringBuilder textBuilder = new StringBuilder();
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(filename, false))//Путь к доку
            {
          
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);
                
                //Загружаем док в поток
                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                //По аттрибутам из стиля оформления вытаскиваем части, которые нас интересуют
                XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
                foreach (XmlNode paragraphNode in paragraphNodes)
                {
                    XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", nsManager);
                    foreach (System.Xml.XmlNode textNode in textNodes)
                    {
                        textBuilder.Append(textNode.InnerText);
                    }
                    textBuilder.Append(Environment.NewLine);
                }

            }
            // По аналогии с прошлым методом, сохраняем всё в json
            string fileNamejSon = @"C:\Users\DIMA\Downloads\Sample3.json";
            JavaScriptSerializer serializer1 = new JavaScriptSerializer();
            string json = serializer1.Serialize(textBuilder.ToString());
            File.WriteAllText(fileNamejSon, json);
            return textBuilder.ToString();
        }

        private static void Main(string[] args)
        {
            string filename = @"C:\Users\DIMA\Downloads\Sample1.docx";//Путь к доку
            var styles = ExtractStylesPart(filename, true);
            var txt = TextFromWord(filename);
            Console.WriteLine(txt.ToString());
            Console.WriteLine(styles.ToString());
            Console.ReadKey();
        }
    }
}
