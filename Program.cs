using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Note;
using System.IO;
using System.Drawing;


namespace OneNoteAnalyzer
{
    class Program
    {
        public static void ExtractAttachment(string onepath, string exportdirectory)
        {
            Console.WriteLine("\n      -> Extracted OneNote Document Attachments: \n");
            string DirectoryName = exportdirectory + "\\OneNoteAttachments";
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            Document OneNoteFile = new Document(onepath);
            IList<AttachedFile> onenotelist = OneNoteFile.GetChildNodes<AttachedFile>();
            foreach (AttachedFile file in onenotelist)
            {
                using (Stream readStream = new MemoryStream(file.Bytes))
                {
                    Console.WriteLine("             -> Extracted Actual Attachment Path: " + Path.GetDirectoryName(file.FilePath) + " | FileName: " + file.FileName + " | Size: " + file.Bytes.Length);
                    using (Stream outStream = System.IO.File.OpenWrite(DirectoryName + "\\" + file.FileName))
                    {
                        
                        readStream.CopyTo(outStream);
                    }
                }
            }
            Console.WriteLine("\n      -> OneNote Document Attachments Extraction Path: " + DirectoryName);

        }

        public static void ExtractMetaData(string onepath)
        {

            Document OneNoteFile = new Document(onepath);
            int pagecount = OneNoteFile.GetChildNodes<Page>().Count;
            Console.WriteLine("\n       -> Page Count: " + pagecount);
            Console.WriteLine("       -> Page MetaData: \n");
            foreach (Page page in OneNoteFile)
            {
                Console.WriteLine("\n       ---------------------------------------------\n");
                Console.WriteLine("             -> Title: " + page.CachedTitleString);
                Console.WriteLine("             -> Author: " + page.Author);
                Console.WriteLine("             -> CreationTime: " + page.CreationTime);
                Console.WriteLine("             -> LastModifiedTime: " + page.LastModifiedTime);
                



            }
            Console.WriteLine("\n       ---------------------------------------------\n");



        }

        public static void ExtractImages(string onepath, string exportdirectory)
        {
            Document OneNoteFile = new Document(onepath);
            string DirectoryName = exportdirectory + "\\OneNoteImages";
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            IList<Aspose.Note.Image> onenodes = OneNoteFile.GetChildNodes<Aspose.Note.Image>();
            Console.WriteLine("\n      -> Extracted OneNote Document Images: \n");
            foreach (Aspose.Note.Image image in onenodes)
            {
                if (image.FileName == null)
                {
                }
                else
                {

                    Console.WriteLine("             -> Extracted Image FileName: " + image.FileName + " | HyperLinkURL: " + (image.HyperlinkUrl == null ? "Null" : image.HyperlinkUrl) + "");
                    using (MemoryStream stream = new MemoryStream(image.Bytes))
                    {
                        using (var filezstream = File.Create(DirectoryName + "\\" + image.FileName))
                        {
                            stream.CopyTo(filezstream);


                        }

                    }

                }


            }
            Console.WriteLine("\n      -> Image Extraction Path: " + DirectoryName);

        }

        public static void ExtractText(string onepath, string exportdirectory)
        {

            Document OneNoteFile = new Document(onepath);
            string DirectoryName = exportdirectory + "\\OneNoteText";
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            Console.WriteLine("\n      -> Extracted OneNote Document Text: \n");
            foreach (Page page in OneNoteFile)
            {
                string pagename = page.CachedTitleString;
                string parsedfilename = string.Join("_", pagename.Split(Path.GetInvalidFileNameChars()));
                string pagepath = DirectoryName + "\\" + parsedfilename + ".txt";
               
                using (StreamWriter textwritten = new StreamWriter(pagepath))
                {
                    string textonenote = string.Join(Environment.NewLine, page.GetChildNodes<RichText>().Select(e => e.Text)) + Environment.NewLine;
                    textwritten.WriteLine(textonenote);
                    Console.WriteLine("             -> Page: " + page.CachedTitleString + " | Extraction Path: " + pagepath);

                }

            }




        }

        public static void ExtractHyperLink(string onepath, string exportdirectory)
        {
            Document OneNoteFile = new Document(onepath);
            string DirectoryName = exportdirectory + "\\OneNoteHyperLinks";
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            string pagepathlink = DirectoryName + "\\onenote_hyperlinks.txt";

            Console.WriteLine("\n      -> Extracted OneNote Document HyperLinks:  (Note: Text might contain hyperlink if no overlay) ");
            foreach (Page page in OneNoteFile)
            {


                IList<RichText> richtextlist = page.GetChildNodes<RichText>();
                string line1 = "\n             -> Page: " + page.CachedTitleString + "\n";
                Console.WriteLine(line1);
                using (StreamWriter linktext = new StreamWriter(pagepathlink,append: true))
                {
                    linktext.WriteLine(line1);
                }

                foreach (RichText Textval in richtextlist)
                {
                    string line2 = "                 -> Text: " + Textval.Text;
                    Console.WriteLine(line2);
                    using (StreamWriter linktext = new StreamWriter(pagepathlink, append: true))
                    {
                        linktext.WriteLine(line2);
                    }
                    foreach (TextStyle style in Textval.Styles)
                    {
                        if (style.HyperlinkAddress == null)
                        {

                        }
                        else
                        {
                            string line3 = "\n                               [*] Contains HyperLink: " + style.HyperlinkAddress + " \n";
                            Console.WriteLine(line3);
                            using (StreamWriter linktext = new StreamWriter(pagepathlink, append: true))
                            {
                                linktext.WriteLine(line3);
                            }
                        }

                    }

                }


            }
            Console.WriteLine("\n      -> HyperLink Extraction Path: " + pagepathlink);
        }

        public static void ConvertToImage(string onepath, string exportdirectory)
        {
            Document OneNoteFile = new Document(onepath);
            string DirectoryName = exportdirectory;
            string FileNameExt = Path.GetFileNameWithoutExtension(onepath);
            string finaldirectory = DirectoryName + "\\ConvertImage_" + FileNameExt + ".png";

            // Save the document as gif.
            OneNoteFile.Save(finaldirectory, SaveFormat.Png);
            Console.WriteLine("\n         -> Saved Path: " + finaldirectory);


        }

        public static string CheckFileFormat(string onepath)
        {

            Document OneNoteFile = new Document(onepath);
            Console.WriteLine("[+] OneNote Document File Format: " + OneNoteFile.FileFormat);
            if (OneNoteFile.FileFormat == FileFormat.Unknown)
            {
                Console.WriteLine("[+] OneNote File Format Not Supported");
                System.Environment.Exit(1);
                return null;

            }
            else
            {
                string filenamewoext = Path.GetFileNameWithoutExtension(onepath);
                string ContentDirectoryName = Path.GetDirectoryName(onepath) + "\\" + filenamewoext + "_content";
                if (!Directory.Exists(ContentDirectoryName))
                {
                    Console.WriteLine("[+] Export Directory Path: " + ContentDirectoryName);
                    Directory.CreateDirectory(ContentDirectoryName);
                    
                }
                return ContentDirectoryName;
            }

        }
        static void Main(string[] args)
        {

            Console.WriteLine(@"
________                 _______          __            _____                .__                              
\_____  \   ____   ____  \      \   _____/  |_  ____   /  _  \   ____ _____  |  | ___.__.________ ___________ 
 /   |   \ /    \_/ __ \ /   |   \ /  _ \   __\/ __ \ /  /_\  \ /    \\__  \ |  |<   |  |\___   // __ \_  __ \
/    |    \   |  \  ___//    |    (  <_> )  | \  ___//    |    \   |  \/ __ \|  |_\___  | /    /\  ___/|  | \/
\_______  /___|  /\___  >____|__  /\____/|__|  \___  >____|__  /___|  (____  /____/ ____|/_____ \\___  >__|   
        \/     \/     \/        \/                 \/        \/     \/     \/     \/           \/    \/       
                                        Author: @knight0x07
                        ");

            if (args == null || args.Length == 0)
            {

                Console.WriteLine("\n[-] Error: No Arguments Passed");
                Console.WriteLine("[-] Usage: OneNoteAnalyzer.exe --file \"<path_to_onenote_document>\"");

            }
            else
            {

                if (args[0] == "--help")
                {
                    Console.WriteLine("\n[-] Usage: OneNoteAnalyzer.exe --file \"<path_to_onenote_document>\"");
                }
                else if (args[0] == "--file")
                {

                    string FilePath = args[1];
                    Console.WriteLine("\n[+] OneNote Document Path: " + FilePath);
                    if (File.Exists(FilePath))
                    {
                        string exportdirectory = CheckFileFormat(FilePath);                       
                        Console.WriteLine("[+] Extracting Attachments from OneNote Document");
                        ExtractAttachment(FilePath, exportdirectory);
                        Console.WriteLine("\n[+] Extracting Page MetaData from OneNote Document");
                        ExtractMetaData(FilePath);
                        Console.WriteLine("\n[+] Extracting Images from OneNote Document");
                        ExtractImages(FilePath, exportdirectory);
                        Console.WriteLine("\n[+] Extracting Text from OneNote Document");
                        ExtractText(FilePath, exportdirectory);
                        Console.WriteLine("\n[+] Extracting HyperLinks from OneNote Document");
                        ExtractHyperLink(FilePath,exportdirectory);
                        Console.WriteLine("\n[+] Converting OneNote Document to Image");
                        ConvertToImage(FilePath,exportdirectory);


                    }
                    else
                    {
                        Console.WriteLine("[-] Eror: Invalid OneNote Document Path");
                    }
                }
                else
                {
                    Console.WriteLine("\n[-] Error: Invalid Arguments");
                    Console.WriteLine("[-] Usage: OneNoteAnalyzer.exe --file \"<path_to_onenote_document>\"");
                }



            }






        }
    }
}
