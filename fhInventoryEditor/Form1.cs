using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Page = UglyToad.PdfPig.Content.Page;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Geometry;
using DataTableAsync;
using System.Net;
using System.Net.Mail;
using System.Runtime.Remoting.Messaging;
using System.Xml.Linq;
using System.Diagnostics.Contracts;

namespace fhInventoryEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static readonly DirectoryInfo jobsFolder = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\");
        private static readonly DirectoryInfo samplesFolder = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\z_samples\\");
        private static readonly string accountName = "Centre Beaubien"; // CentreLeCap, Centre Beaubien
        private static readonly DirectoryInfo jobInfo = new DirectoryInfo($"{jobsFolder.FullName}{accountName}\\");
        private static readonly string htmlPath = $"{jobInfo.FullName}a_jobSummary\\index.html";
        private readonly HtmlDocument htmlEditor = new HtmlDocument();

        public enum PageRegion { none, contact, table_disclaimer, table_hdr, table_data, footer }
        public enum DocumentType { none, quote, delivery, invoice }
        public enum DocumentLanguage { none, english, french }
        public struct Document
        {
            public DocumentType type;
            public DocumentLanguage language;
            public override string ToString() => $"{type} [{language}]";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /// objective: go into any folder and extract contact details (po#, client name address, etc) + list of products
            /// 
            /// SaveAll_byType()
            ///     1] uses Get_DocumentTypes() to determine the document type and language, but does not parse it
            ///     2] determines how the file should be named and saves if it doesn't use the set naming convention
            ///     3] returns a dictionary<string, string> where:
            ///         a) the key is the existing file name, correct or incorrect convention
            ///         b) the value is only populated with the correct convention if the old name didn't follow the convention 
            /// Get_DocumentTypes()
            ///     1] opens the pdf file and determines the document type and language, but does not parse it
            ///     2] groups pdf files, by folder- or all folders into a dictionary<DocumentType, List<FileInfo>>

            //var type = Get_documentTypeLanguage();
            //Debugger.Break();
            //Get_samples();
            //Debugger.Break();
            // new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Centre Evasion\\")

            DateTime startTime = DateTime.Now;
            var forms = Parse_forms();
            string contacts = string.Join(Environment.NewLine, forms.Item2.Keys);
            string html = forms.Item1.HTML;
            //var moves = SaveAll_byType("Centre Evasion");
            DateTime endTime = DateTime.Now;
            TimeSpan elapsed = endTime - startTime;
            Debugger.Break();

            startTime = DateTime.Now;
            //SaveAll_byType();
            //var doctypes = Get_DocumentTypes("Centre Evasion");
            //var form = Parse_form(new FileInfo(moves.Skip(2).First().Key));
            endTime = DateTime.Now;
            elapsed = endTime - startTime;
            Debugger.Break();

            //const string exDlvry = "C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Cadens Lighthouse\\delivery_Cadens Lighthouse.pdf";
            //const string exQuote = "C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Centre Evasion\\quote_Centre Evasion.pdf";
            //bool exIsQuote = true;
            //var yy = Parse_form(new FileInfo(exIsQuote ? exQuote : exDlvry));
            //Debugger.Break();

            var tables = new Dictionary<string, string>();
            var jobs = new List<DirectoryInfo>(new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\").EnumerateDirectories());
            jobs = new List<DirectoryInfo>(jobs.Where(d => !(d.Name.StartsWith("z_") | d.Name.StartsWith("a)"))));
            foreach (var job in jobs)
            {
                var tableFiles = new List<FileInfo>(job.EnumerateFiles("*.pdf", SearchOption.TopDirectoryOnly));
                var dlvryFiles = new List<FileInfo>(tableFiles.Where(f => f.Name.StartsWith("delivery")));
                var quoteFiles = new List<FileInfo>(tableFiles.Where(f => f.Name.StartsWith("quote")));
                foreach (var dlvryFile in dlvryFiles.Take(100))
                {
                    try
                    {
                        var xx = Parse_form(dlvryFile, false);
                        tables.Add(dlvryFile.Name, xx.Item1.HTML);
                    }
                    catch { }
                }
                foreach (var quoteFile in quoteFiles)
                {
                    try
                    {
                        var xx = Parse_form(quoteFile, false);
                        tables.Add(quoteFile.Name, xx.Item1.HTML);
                    }
                    catch { Debugger.Break(); }
                }
            }
            Debugger.Break();

            htmlEditor.LoadHtml(File.ReadAllText(htmlPath));
            //htmlEditor.GetElementbyId("businessName").SetAttributeValue("value", "Centre Le Cap");
            //htmlEditor.GetElementbyId("purchaseOrder").SetAttributeValue("value", "");
            //using (StreamWriter sw = new StreamWriter(htmlPath))
            //{
            //    htmlEditor.Save(sw);
            //}
        }
        private static void Get_samples()
        {
            var samples = Get_DocumentTypes();
            samples.Remove(DocumentType.none);
            foreach (var sampleType in samples)
                foreach (var pdf in sampleType.Value)
                    File.Copy(pdf.FullName, $"{samplesFolder}{pdf.Name}", true);
        }
        private static DirectoryInfo Get_directoryByName(string foldername_Or_Filename)
        {
            if (foldername_Or_Filename == null) return null;
            var matchedDirectories = new List<DirectoryInfo>();
            if (foldername_Or_Filename.EndsWith(".txt") | foldername_Or_Filename.EndsWith(".pdf"))
            {
                string filetype = foldername_Or_Filename.Split('.').Last();
                var allFiles = new List<FileInfo>(jobsFolder.EnumerateFiles($"*.{filetype}", SearchOption.AllDirectories));
                foreach (var file in allFiles) if (file.Name == foldername_Or_Filename) matchedDirectories.Add(file.Directory);
            }
            else
                matchedDirectories.AddRange(jobsFolder.EnumerateDirectories().Where(d => d.Name == foldername_Or_Filename));

            if (matchedDirectories.Any())
                return matchedDirectories.First();
            return null;
        }
        internal static Dictionary<DocumentType, List<FileInfo>> Get_DocumentTypes(string foldername_Or_Filename)
        {
            DirectoryInfo directory = Get_directoryByName(foldername_Or_Filename);
            return directory == null ? null : Get_DocumentTypes(directory);
        }
        private static Dictionary<DocumentType, List<FileInfo>> Get_DocumentTypes()
        {
            var doctypes = new Dictionary<DocumentType, List<FileInfo>>();
            foreach (var doctype in Enum.GetNames(typeof(DocumentType)))
                doctypes[(DocumentType)Enum.Parse(typeof(DocumentType), doctype)] = new List<FileInfo>();
            foreach (var jobFolder in jobsFolder.EnumerateDirectories().Where(d => !(d.Name.StartsWith("a)") | d.Name.StartsWith("z_"))))
            {
                foreach (var doctypeList in Get_DocumentTypes(jobFolder))
                    doctypes[doctypeList.Key].AddRange(doctypeList.Value);
            }
            return doctypes;
        }
        internal static Dictionary<DocumentType, List<FileInfo>> Get_DocumentTypes(DirectoryInfo jobfolder)
        {
            List<FileInfo> jobFiles = new List<FileInfo>(jobfolder.EnumerateFiles("*.pdf", SearchOption.AllDirectories));
            var doctypes = new Dictionary<DocumentType, List<FileInfo>>();
            foreach (var pdf in jobFiles)
            {
                var pdfDocument = Get_documentTypeLanguage(pdf);
                if (!doctypes.ContainsKey(pdfDocument.type)) doctypes[pdfDocument.type] = new List<FileInfo>();
                doctypes[pdfDocument.type].Add(pdf);
            }
            return doctypes;
        }
        private static Dictionary<DocumentType, List<FileInfo>> Get_DocumentTypes(FileInfo jobinfo) => Get_DocumentTypes(jobinfo.Directory);
        internal static Dictionary<string, string> SaveAll_byType()
        {
            var allMoves = new Dictionary<string, string>();
            var allFolders = new List<DirectoryInfo>(jobsFolder.EnumerateDirectories());
            foreach (var folder in allFolders)
                foreach (var filePair in SaveAll_byType(folder))
                    allMoves.Add(filePair.Key, filePair.Value);
            return allMoves;
        }
        internal static Dictionary<string, string> SaveAll_byType(string foldername_Or_Filename) => SaveAll_byType(Get_directoryByName(foldername_Or_Filename));
        private static Dictionary<string, string> SaveAll_byType(DirectoryInfo jobFolder) => SaveAll_byDocType(Get_DocumentTypes(jobFolder));
        internal static Dictionary<string, string> SaveAll_byType(FileInfo jobinfo) => SaveAll_byDocType(Get_DocumentTypes(jobinfo));
        private static Dictionary<string, string> SaveAll_byDocType(Dictionary<DocumentType, List<FileInfo>> pdfTypes)
        {
            if (pdfTypes.ContainsKey(DocumentType.none)) pdfTypes.Remove(DocumentType.none);
            var moves = new Dictionary<string, string>();
            foreach (var pdfType in pdfTypes)
            {
                byte fileIndex = 0;
                foreach (var pdf in pdfType.Value)
                {
                    string oldPath = pdf.FullName;
                    string newPath = string.Empty;
                    string[] splitLevels = oldPath.Split('\\');
                    byte lvlIndex = 0;
                    string lvlLast = string.Empty;
                    foreach (string lvl in splitLevels)
                    {
                        newPath += lvl + '\\';
                        if (lvlLast != lvl)
                        {
                            if (lvlLast.ToLowerInvariant() == "jobs")
                            {
                                if (pdfType.Value.Count == 1) newPath += $"{pdfType.Key}_{lvl}.pdf";
                                else newPath += $"{pdfType.Key}_{lvl} [{fileIndex}].pdf";
                                break;
                            }
                            lvlLast = lvl;
                        }
                        lvlIndex++;
                    }
                    fileIndex++;
                    moves.Add(oldPath, string.Empty);
                    if (newPath != oldPath)
                    {
                        string guidPath = $"{pdf.Directory.FullName}" + '\\' + Guid.NewGuid().ToString() + ".pdf";
                        File.Move(oldPath, guidPath);
                        File.Move(guidPath, newPath);
                        moves[oldPath] = newPath;
                    }
                }
            }
            return moves;
        }
        private static Document Get_documentTypeLanguage(FileInfo pdfinfo)
        {
            string[] keyDeliveryWords = new string[]
            {
                "Formulaire de vérification de livraison",
                "Les commandes doivent être inspectées pour dommages d'expédition",
                "Shipping damage claims will only be accepted by",
                "Delivery Verification Form"
            };
            var documentPages = Get_pdfText(pdfinfo);
            if (documentPages == null) return new Document() { language = DocumentLanguage.none, type = DocumentType.none };
            var documentText = string.Join(Environment.NewLine, documentPages.Values);
            
            // quote match MUST be first as the quote contains the delivery disclaimer statement, but delivery forms DON'T have QUOTE#/SOUMISSION
            var quoteMatch = Regex.Match(documentText, "(QUOTE|SOUMISSION) {0,}#:", RegexOptions.IgnoreCase);
            if (quoteMatch.Success)
            {
                var quoteLanguage = Regex.IsMatch(quoteMatch.Value, "SOUMISSION {0,1}#:") ? DocumentLanguage.french : DocumentLanguage.english;
                return new Document() { language = quoteLanguage, type = DocumentType.quote };
            }
            
            var dlvryMatch = Regex.Match(documentText, $"({string.Join("|", keyDeliveryWords)})", RegexOptions.IgnoreCase);
            if (dlvryMatch.Success)
            {
                var deliveryLanguage = Regex.IsMatch(dlvryMatch.Value, $"({string.Join("|", keyDeliveryWords.Take(2))})") ? DocumentLanguage.french : DocumentLanguage.english;
                return new Document() { language = deliveryLanguage, type = DocumentType.delivery };
            }

            var noneLanguage = Regex.IsMatch(documentText, "[àâçéèêëîïôûùüÿñæœ]", RegexOptions.IgnoreCase) ? DocumentLanguage.french : DocumentLanguage.english;
            return new Document() { language = noneLanguage, type = DocumentType.none };
        }
        private static Dictionary<int, string> Get_pdfText(FileInfo pdfinfo)
        {
            if (pdfinfo == null || !File.Exists(pdfinfo.FullName)) { return null; }
            try
            {
                Dictionary<int, string> pages = new Dictionary<int, string>();
                using (PdfDocument document = PdfDocument.Open(pdfinfo.FullName))
                    foreach (var page in document.GetPages())
                        pages[page.Number] = page.Text;
                return pages;
            }
            catch { return null; }
        }
        private static string CleanDescription(string descriptionIn)
        {
            string pad = new string(' ', 8);
            descriptionIn += pad;
            string newDescription = descriptionIn;
            newDescription = Regex.Replace(newDescription, "\\sTRK\\s", "     ") + pad; // dont know what these are
            newDescription = Regex.Replace(newDescription, " N[A-Z][0-9]{2} ", pad) + pad; // NS15 or NF13 etc
            while (Regex.IsMatch(newDescription, " ([A-Z])\\1 "))
                newDescription = Regex.Replace(newDescription, " ([A-Z])\\1 ", pad) + pad; // AA, BB, etc
            newDescription = Regex.Replace(newDescription, " {2,}", " ") + pad;
            newDescription = newDescription.Trim();
            return newDescription;
        }
        private static Dictionary<bool, List<Word>> Words_inLine(PdfRectangle rectIn, List<Word> words, byte margin = 0)
        {
            var xxx = new PdfRectangle(new PdfPoint(0, rectIn.BottomLeft.Y - margin), new PdfPoint(1000, rectIn.TopRight.Y + margin));
            var wordsInRects = new List<Word>(words.Where(w => w.BoundingBox.IntersectsWith(xxx)));
            var words_inLine = new Dictionary<bool, List<Word>>();
            foreach (var word in wordsInRects)
            {
                bool isAbove = word.BoundingBox.Top > xxx.Centroid.Y;
                if (!words_inLine.ContainsKey(isAbove)) words_inLine[isAbove] = new List<Word>();
                words_inLine[isAbove].Add(word);
            }
            foreach (var linegrp in words_inLine)
                linegrp.Value.Sort((w1, w2) => w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left));
            return words_inLine;
        }
        internal static Tuple<Table, Dictionary<string, List<string>>> Parse_forms(bool openFiles = false)
        {
            Table allTables = new Table();
            Dictionary<string, List<string>> allContacts = new Dictionary<string, List<string>>();
            foreach (var jobFolder in jobsFolder.EnumerateDirectories())
            {
                var parsedForms = Parse_forms(jobFolder, openFiles);
                allTables.Merge(parsedForms.Item1);
                foreach (var contact in parsedForms.Item2)
                {
                    if (!allContacts.ContainsKey(contact.Key)) allContacts[contact.Key] = new List<string>();
                    allContacts[contact.Key].AddRange(contact.Value);
                }
            }
            return Tuple.Create(allTables, allContacts);
        }
        private static Tuple<Table, Dictionary<string, List<string>>> Parse_forms(DirectoryInfo jobFolder, bool openFiles = false)
        {
            Table folderTable = new Table();
            Dictionary<string, List<string>> contacts = new Dictionary<string, List<string>>();
            foreach (var file in jobFolder.EnumerateFiles("*.pdf").Where(f=>Regex.IsMatch(f.Name, "(delivery|quote)_[^.]{1,}\\.pdf")))
            {
                var parsedForm = Parse_form(file, openFiles);
                folderTable.Merge(parsedForm.Item1);
                foreach (var contact in parsedForm.Item3)
                {
                    if (!contacts.ContainsKey(contact.Key)) contacts[contact.Key] = new List<string>();
                    contacts[contact.Key].Add(contact.Value);
                }
            }
            return Tuple.Create(folderTable, contacts);
        }
        private static Tuple<Table, Document, Dictionary<string, string>> Parse_form(FileInfo jobinfo, bool openFile = false)
        {
            if (jobinfo == null) return Tuple.Create(new Table(), new Document(), new Dictionary<string, string>());
            else
            {
                if (openFile) Process.Start($"{jobinfo.FullName}");

                var ColumnNames = new List<string>();

                Table itemTable = new Table();
                var item = itemTable.Columns.Add("item", typeof(string));
                var desc = itemTable.Columns.Add("desc", typeof(string));
                var qty = itemTable.Columns.Add("qty", typeof(double));
                itemTable.PrimaryKeys = new Table.Column[] { item };

                var contacts = new Dictionary<string, string>();
                var tableRows = new Dictionary<byte, string>();
                var documentTypeLanguage = Get_documentTypeLanguage(jobinfo);

                using (PdfDocument document = PdfDocument.Open(jobinfo.FullName))
                {
                    var fonts_byPage = new Dictionary<byte, Dictionary<string, List<Word>>>();
                    var letters_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<double, Dictionary<int, Letter>>>>();
                    var consecutiveLetters_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<byte, Dictionary<int, Letter>>>>();
                    var colRects = new Dictionary<string, PdfRectangle>();
                    var colNames = new Dictionary<string, Word>();
                    var cols = new List<double>();

                    foreach (Page page in document.GetPages())
                    {
                        string pageText = page.Text;
                        PageRegion pgRegion = new PageRegion();

                        var boxes = page.ExperimentalAccess.Paths.Select(p => p.GetBoundingRectangle()).Where(bb => bb.HasValue).Select(bb => bb.Value).ToList();
                        var pdfRects = new List<PdfRectangle>(boxes.Distinct());
                        pdfRects.Sort((r1, r2) =>
                        {
                            int lvl1 = r2.Bottom.CompareTo(r1.Bottom);
                            if (lvl1 != 0)
                                return lvl1;
                            else
                            {
                                int lvl2 = r1.Width.CompareTo(r2.Width);
                                if (lvl2 != 0) { return lvl2; }
                                else
                                {
                                    int lvl3 = r1.Left.CompareTo(r2.Left);
                                    return lvl3;
                                }
                            }
                        });

                        List<Word> words = new List<Word>(page.GetWords());
                        words.Sort((w1, w2) =>
                        {
                            int lvl1 = w2.BoundingBox.Bottom.CompareTo(w1.BoundingBox.Bottom);
                            if (lvl1 != 0) return lvl1;
                            else
                            {
                                int lvl2 = w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left);
                                return lvl2;
                            }
                        });

                        List<Letter> letters = new List<Letter>(page.Letters.OrderBy(l => l.GlyphRectangle.Left));

                        byte pageNbr = Convert.ToByte(page.Number);
                        if (Regex.IsMatch(pageText, "ITEM {0,}# {0,}DESCRIPTION {0,}Q", RegexOptions.IgnoreCase))
                        {
                            var widthDictionary = new Dictionary<double, List<PdfRectangle>>();
                            var pdfRects_exceptText = new List<PdfRectangle>();
                            var pdfRects_Lines = new List<PdfRectangle>();
                            var pdfRects_linesHorizontal = new List<PdfRectangle>();
                            var pdfRects_linesVertical = new List<PdfRectangle>();
                            double thinnestVerticalLine = 99;
                            foreach (PdfRectangle r1 in pdfRects)
                            {
                                if (!widthDictionary.ContainsKey(r1.Width)) widthDictionary[r1.Width] = new List<PdfRectangle>();
                                widthDictionary[r1.Width].Add(r1);
                                bool foundWordRect = false;
                                bool foundWordInRect = false;
                                foreach (Word w1 in words)
                                {
                                    var r2 = w1.BoundingBox;
                                    if (r1.Area == r2.Area & r1.Left == r2.Left & r1.Top == r2.Top & r1.Bottom == r2.Bottom & r1.Right == r2.Right)
                                    {
                                        foundWordRect = true;
                                        break;
                                    }
                                    if (r1.Contains(r2))
                                    {
                                        foundWordInRect = true;
                                        break;
                                    }
                                }
                                bool foundLetterRect = false;
                                bool foundLetterInRect = false;
                                foreach (Letter l in letters)
                                {
                                    var r2 = l.GlyphRectangle;
                                    if (r1.Area == r2.Area & r1.Left == r2.Left & r1.Top == r2.Top & r1.Bottom == r2.Bottom & r1.Right == r2.Right)
                                    {
                                        foundLetterRect = true;
                                        break;
                                    }
                                    if (r1.Contains(r2))
                                    {
                                        foundLetterInRect = true;
                                        break;
                                    }
                                }
                                if (!(foundWordRect | foundLetterRect)) pdfRects_exceptText.Add(r1);
                                if (!(foundWordInRect | foundLetterInRect)) pdfRects_Lines.Add(r1);
                                if (r1.Height >= 0 & r1.Height < 2) pdfRects_linesHorizontal.Add(r1);
                                if (r1.Width >= 0 & r1.Width < 2)
                                {
                                    pdfRects_linesVertical.Add(r1);
                                    if (thinnestVerticalLine > r1.Width) thinnestVerticalLine = r1.Width;
                                } // MUST BE > AND = 0! some lines come through with 0 width
                            }
                            thinnestVerticalLine = Math.Round(thinnestVerticalLine, 2);
                            pdfRects_linesHorizontal.Sort((l1, l2) =>
                            {
                                int lvl1 = l2.Bottom.CompareTo(l1.Bottom);
                                if (lvl1 != 0) return lvl1;
                                else
                                {
                                    int lvl2 = l2.Width.CompareTo(l1.Width);
                                    return lvl2;
                                }
                            });
                            pdfRects_linesVertical.Sort((l1, l2) =>
                            {
                                int lvl1 = l1.Left.CompareTo(l2.Left);
                                if (lvl1 != 0) return lvl1;
                                else
                                {
                                    int lvl2 = l2.Height.CompareTo(l1.Height);
                                    return lvl2;
                                }
                            });
                            widthDictionary = widthDictionary.OrderByDescending(w => w.Value.Count()).ToDictionary(k => k.Key, v => v.Value);
                            double tableWidth = widthDictionary.Keys.First();
                            var tableWidths = new List<PdfRectangle>(pdfRects.Where(r => r.Width == tableWidth).OrderByDescending(r => r.Area));
                            var table_linesVertical = new Dictionary<double, List<PdfRectangle>>();
                            foreach (var vLine in pdfRects_linesVertical.Where(l => Math.Round(l.Width, 2) == thinnestVerticalLine))
                            {
                                if (!table_linesVertical.ContainsKey(vLine.Left)) table_linesVertical[vLine.Left] = new List<PdfRectangle>();
                                table_linesVertical[vLine.Left].Add(vLine);
                            }
                            var tableColumnLefts = new Dictionary<int, PdfRectangle>();
                            foreach (var table_lineVertical in table_linesVertical)
                                tableColumnLefts.Add(tableColumnLefts.Count, table_lineVertical.Value.First());

                            /// will be exactly 6 vertical lines... less or more means something is wrong
                            // [0] {[(x:50.664, y:560.57), 0.719999999999999, 434.28]} left side of table / column ITEM#
                            // [1] {[(x:100.37, y:560.57), 0.719989999999996, 434.28]} left side of column DESCRIPTION
                            // [2] {[(x:269.14, y:560.57), 0.719969999999989, 434.28]} left side of column QTY
                            // [3] {[(x:314.52, y:560.57), 0.720000000000027, 434.28]} left side of column RECEIVED
                            // [4] {[(x:362.54, y:560.57), 0.720000000000027, 434.28]} left side of column MISSING
                            // [5] {[(x:405.5, y:560.57), 0.720000000000027, 434.28]} right side of table / column MISSING

                            double col1_itemLeft = tableColumnLefts[0].Left;
                            double col2_descLeft = tableColumnLefts[1].Left;
                            double col3_qtyLeft = tableColumnLefts[2].Left;

                            double tableDataTop = tableColumnLefts[1].Top; // column[0] line may extend all the way to the top of the document so use 1 which should stop at the table top
                            var wordsAboveTable = new List<Word>(words.Where(w => w.BoundingBox.Bottom > tableDataTop));
                            if (documentTypeLanguage.type == DocumentType.delivery)
                            {
                                double VERIFICATION = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "v(e|é)rification", RegexOptions.IgnoreCase))).FirstOrDefault().BoundingBox.Bottom;
                                double PLEASENOTE = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "NOTE(R){0,1}:"))).FirstOrDefault().BoundingBox.Top;
                                var wordsAboveDisclaimer = new List<Word>(wordsAboveTable.Where(w => w.BoundingBox.Bottom > PLEASENOTE));
                                var wordsContact = new List<Word>(wordsAboveDisclaimer.Where(w => w.BoundingBox.Top < VERIFICATION));
                                contacts = Get_contactKeysValues(wordsContact, col3_qtyLeft, 3);
                            }
                            if (documentTypeLanguage.type == DocumentType.quote)
                            {
                                var phoneNumber = new List<Word>(words.Where(w => w.Text.Contains("1-800-265-6900")));
                                if (phoneNumber.Any())
                                {
                                    var wordsContact = new List<Word>(wordsAboveTable.Where(w => w.BoundingBox.Top < phoneNumber.First().BoundingBox.Bottom));
                                    contacts = Get_contactKeysValues(wordsContact, col3_qtyLeft, 0);
                                }
                            }

                            #region" column names "
                            var firstColumnName = new List<Word>(words.Where(w => w.Text == "ITEM"));
                            if (firstColumnName.Any())
                            {
                                var itemNbr = firstColumnName.First();
                                var allColumnNames = new List<Word>(words.Where(w => w.BoundingBox.Bottom == itemNbr.BoundingBox.Bottom));
                                ColumnNames = new List<string>(allColumnNames.Where(c => c.Text != "#").Select(c => c.Text));

                                colNames = words.Where(w => ColumnNames.Contains(w.Text.Trim())).ToDictionary(k => k.Text, v => v);
                                colRects = colNames.ToDictionary(k => k.Key, v => v.Value.BoundingBox);
                            }
                            #endregion
                        }
                        if (words.Any())
                        {
                            #region" fonts --> size(w,h) "
                            var fonts = new Dictionary<string, List<Word>>();
                            foreach (var word in words)
                            {
                                string fontname = word.FontName;
                                if (!fonts.ContainsKey(fontname)) fonts[fontname] = new List<Word>();
                                fonts[fontname].Add(word);
                            }
                            fonts_byPage[pageNbr] = fonts;
                            Dictionary<string, double> width_byFont = new Dictionary<string, double>();
                            Dictionary<string, double> height_byFont = new Dictionary<string, double>();
                            foreach (string fontName in fonts.Keys)
                            {
                                double fw = 0;
                                double fh = 0;
                                List<Word> fontWords = fonts[fontName];
                                fontWords.Sort((y1, y2) => y2.Letters.Count.CompareTo(y1.Letters.Count));
                                if (fontWords.Any())
                                {
                                    List<Letter> longestWordLetters = new List<Letter>(fontWords.First().Letters);
                                    List<double> fontXs = new List<double>(longestWordLetters.Select(x => Math.Round(x.Location.X, 1)).Distinct());
                                    List<double> fontHs = new List<double>(longestWordLetters.Select(x => Math.Round(x.GlyphRectangle.Height, 1)).Distinct());
                                    fw = Math.Round((fontXs.Max() - fontXs.Min()) / (fontXs.Count - 1), 1);
                                    fh = Math.Round(fontHs.Max(), 1);
                                }
                                width_byFont.Add(fontName, fw);
                                height_byFont.Add(fontName, fh);
                            }
                            fonts = fonts.OrderByDescending(f => f.Value.Count).ToDictionary(k => k.Key, v => v.Value);
                            string mostCommonFont = fonts.Keys.FirstOrDefault().ToString();
                            double fontWidth = width_byFont[mostCommonFont];
                            double fontHeight = 10; //height_byFont[mostCommonFont];
                            #endregion
                            #region" letters_thisPage "
                            var letters_thisPage = new Dictionary<byte, Dictionary<double, Dictionary<int, Letter>>>();
                            foreach (var ltr in page.Letters)
                            {
                                double Y = ltr.StartBaseLine.Y;
                                byte lineNbr = Convert.ToByte(Y / fontHeight);
                                if (!letters_thisPage.ContainsKey(lineNbr)) letters_thisPage[lineNbr] = new Dictionary<double, Dictionary<int, Letter>>();
                                if (!letters_thisPage[lineNbr].ContainsKey(Y)) letters_thisPage[lineNbr][Y] = new Dictionary<int, Letter>();
                                letters_thisPage[lineNbr][Y].Add(letters_thisPage[lineNbr][Y].Count, ltr);
                            }
                            letters_thisPage = letters_thisPage.OrderByDescending(y => y.Key).ToDictionary(k => k.Key, y => y.Value);
                            letters_byPage[pageNbr] = letters_thisPage;
                            #endregion
                            var lines_thisPage = new Dictionary<double, string>();
                            var lines = new List<string>();
                            consecutiveLetters_byPage[pageNbr] = new Dictionary<byte, Dictionary<byte, Dictionary<int, Letter>>>();
                            foreach (byte lineNbr in letters_thisPage.Keys)
                            {
                                byte groupIndex = 0;
                                var consecutiveLetters = new Dictionary<byte, Dictionary<int, Letter>>();
                                var lineGroups = letters_thisPage[lineNbr];
                                foreach (double lineY in lineGroups.Keys)
                                {
                                    if (lineY != 0)
                                    {
                                        Dictionary<int, Letter> indexedLetters = letters_thisPage[lineNbr][lineY];
                                        Dictionary<int, Letter> remainingLetters = new Dictionary<int, Letter>(indexedLetters);
                                        KeyValuePair<int, Letter> firstInChain = new KeyValuePair<int, Letter>(0, indexedLetters[0]);
                                        consecutiveLetters[groupIndex] = new Dictionary<int, Letter>();

                                        while (remainingLetters.Any())
                                        {
                                            remainingLetters.Remove(firstInChain.Key);
                                            consecutiveLetters[groupIndex].Add(consecutiveLetters[groupIndex].Count, firstInChain.Value);
                                            var nextInChain = new List<KeyValuePair<int, Letter>>(remainingLetters.Where(nextLetter => (nextLetter.Value.StartBaseLine.X - firstInChain.Value.EndBaseLine.X) < 2));
                                            if (nextInChain.Any()) { firstInChain = nextInChain.First(); }
                                            else if (remainingLetters.Any())
                                            {
                                                firstInChain = remainingLetters.First();
                                                groupIndex++;
                                                consecutiveLetters[groupIndex] = new Dictionary<int, Letter>();
                                            }
                                        }
                                        groupIndex++;
                                    }
                                    groupIndex++;
                                }
                                var letterStrings = new Dictionary<byte, Dictionary<int, Letter>>(consecutiveLetters);
                                consecutiveLetters.Clear();

                                foreach (var letterGroup in letterStrings.OrderBy(cl => cl.Value.Min(l => l.Value.StartBaseLine.X)))
                                    consecutiveLetters.Add((byte)consecutiveLetters.Count, letterGroup.Value);

                                consecutiveLetters_byPage[pageNbr][lineNbr] = consecutiveLetters;
                                string line = string.Join("■", consecutiveLetters.Select(cl => string.Join(string.Empty, cl.Value.Select(l => l.Value.Value))));
                                string lineData = $"{lineNbr:000}_{consecutiveLetters[0][0].StartBaseLine.X:000.0}|{line}";
                                lines.Add(lineData);
                                lines_thisPage[lineNbr] = line;
                                if (lineData.Contains("ITEM #"))
                                {
                                    var colWord = colNames["ITEM"];
                                    var colRect = colWord.BoundingBox;
                                    var boundingRects = new List<PdfRectangle>(pdfRects.Where(r => r.Contains(colRect)).OrderBy(r => r.Area)); // smallest to largest
                                    var tableHeadRect = boundingRects.FirstOrDefault(); // this should be the table heading rectangle that contains the column names
                                    var intersectRects = new List<PdfRectangle>(pdfRects.Where(r => tableHeadRect.IntersectsWith(r)).OrderBy(r => r.Width));
                                    foreach (PdfRectangle r1 in intersectRects)
                                    {
                                        foreach (PdfRectangle r2 in colRects.Values)
                                        {
                                            if (r1.Area == r2.Area & r1.Left == r2.Left & r1.Top == r2.Top & r1.Bottom == r2.Bottom & r1.Right == r2.Right)
                                                Debugger.Break(); // dont want any of the colrects in the intersects results (and cant use Except)
                                        }
                                        foreach (Word w in words)
                                        {
                                            var r2 = w.BoundingBox;
                                            if (r1.Area == r2.Area & r1.Left == r2.Left & r1.Top == r2.Top & r1.Bottom == r2.Bottom & r1.Right == r2.Right)
                                                Debugger.Break(); // dont want any of the colrects in the intersects results (and cant use Except)
                                        }
                                        foreach (Letter l in letters)
                                        {
                                            var r2 = l.GlyphRectangle;
                                            if (r1.Area == r2.Area & r1.Left == r2.Left & r1.Top == r2.Top & r1.Bottom == r2.Bottom & r1.Right == r2.Right)
                                                Debugger.Break(); // dont want any of the colrects in the intersects results (and cant use Except)
                                        }
                                    }
                                    var lefts = new List<double>(intersectRects.Select(r => r.Left).Distinct());
                                    lefts.Sort();
                                    foreach (double l1 in lefts)
                                    {
                                        var leftGroup = new List<double>();
                                        foreach (var l2 in lefts)
                                        {
                                            double min = new double[] { l1, l2 }.Min();
                                            double max = new double[] { l1, l2 }.Max();
                                            double min_max = min / max;
                                            if (min_max >= .9) leftGroup.Add(l2);
                                        }
                                        if (leftGroup.Any()) cols.Add(leftGroup.Min());
                                    }
                                    cols = cols.Distinct().ToList();
                                    cols.Sort();
                                    var cols_colRects = new List<double>();
                                    foreach (var cRect in colRects)
                                    {
                                        var cR = cols.Where(c => c < cRect.Value.Left).Max();
                                        cols_colRects.Add(cR);
                                    }
                                    cols = cols_colRects;
                                    cols.Sort();

                                    pgRegion = PageRegion.table_data;
                                }
                                else if (pgRegion == PageRegion.table_data)
                                {
                                    const byte wordGrp = 0;
                                    var letterGrp = consecutiveLetters[wordGrp];
                                    Letter firstLetter_inWord = letterGrp[0];
                                    var colIndex = cols.IndexOf(cols.Where(c => c < firstLetter_inWord.StartBaseLine.X).Max());
                                    string code = string.Join(string.Empty, letterGrp.OrderBy(l => l.Key).Select(l => l.Value.Value)).Replace(" #", string.Empty).Trim();
                                    if (colIndex == 0)
                                    {
                                        PdfRectangle letterGlyph = firstLetter_inWord.GlyphRectangle;
                                        var words_thisLine = Words_inLine(letterGlyph, words);
                                        var words_byColumn = new Dictionary<string, List<Word>>();
                                        foreach (var isAbove in words_thisLine)
                                        {
                                            foreach (var word in words_thisLine[isAbove.Key])
                                            {
                                                colIndex = cols.IndexOf(cols.Where(c => c < word.Letters[0].StartBaseLine.X).Max());
                                                string colName = ColumnNames[colIndex];
                                                if (!words_byColumn.ContainsKey(colName)) words_byColumn[colName] = new List<Word>();
                                                words_byColumn[colName].Add(word);
                                            }
                                        }
                                        var columnWords = new Dictionary<string, string>();
                                        foreach (var col in words_byColumn)
                                        {
                                            var cleanDescription = CleanDescription(string.Join(" ", col.Value.Select(w => w.Text)));
                                            columnWords[col.Key] = cleanDescription;
                                            //if (cleanDescription.Contains("ULTRA")) Debugger.Break();
                                        }
                                        foreach (var columnName in ColumnNames) if (!columnWords.ContainsKey(columnName)) columnWords[columnName] = string.Empty;
                                        tableRows.Add((byte)tableRows.Count, JsonConvert.SerializeObject(columnWords, Formatting.None));

                                        string cell1_item = columnWords[ColumnNames[0]];
                                        string cell2_desc = columnWords[ColumnNames[1]];
                                        string cell3_qty = columnWords.ContainsKey(ColumnNames[2]) ? columnWords[ColumnNames[2]] : string.Empty; // may not contain (ex. MILKY WAY CARPET KIT CONSISTS OF)
                                        string cell3_nbrs = Regex.Match(cell3_qty, "[$0-9,.]{1,}").Value;
                                        double.TryParse(cell3_nbrs, out double qtyCell);
                                        itemTable.Rows.Add(new object[] { cell1_item, cell2_desc, qtyCell });
                                    } // get all the words in a table row once (for column 0)
                                }

                                if (pgRegion == PageRegion.none & Regex.IsMatch(line, "v(e|é)rification", RegexOptions.IgnoreCase))
                                    pgRegion = PageRegion.contact;
                                else if (pgRegion == PageRegion.contact) { }
                                else if (pgRegion == PageRegion.table_data)
                                {
                                    // [117.5] SIGNATURE ■[336.8] DATE
                                    bool isTableEnd = line.Contains("SIGNATURE") & line.Contains("DATE");
                                    if (line.Contains("NOM IMPRIMÉ") | line.Contains("PRINTED NAME")) isTableEnd = true;
                                    if (line.Contains("VEUILLEZ ENVOYER")) isTableEnd = true;
                                    if (line.Contains("ONCE COMPLETED")) isTableEnd = true;
                                    if (Regex.IsMatch(line, "page [0-9] of [0-9]", RegexOptions.IgnoreCase)) isTableEnd = true;
                                    if (line.Contains("TERMES ET CONDITIONS")) isTableEnd = true;
                                    if (line.Contains("TERMS AND CONDITIONS")) isTableEnd = true;
                                    if (line.Contains("VIRTUAL INSTALLATION AVAILABLE")) isTableEnd = true;
                                    if (!line.Any()) isTableEnd = true;
                                    if (isTableEnd)
                                        pgRegion = PageRegion.footer;
                                }
                            }
                        }
                    }

                    #region" save .txt file - MUST assume pdf filename is the correct format "
                    // save the contacts dictionary + products table in a Tuple as -->  delivery_Centre Evasion
                    string directoryFullPath = $"{jobinfo.DirectoryName}\\{jobinfo.Directory.Name}";
                    var newFilePath = new FileInfo(jobinfo.FullName.Replace(".pdf", ".txt"));
                    var fileTuple = Tuple.Create(contacts, itemTable);
                    File.WriteAllText(newFilePath.FullName, JsonConvert.SerializeObject(fileTuple, Formatting.Indented));
                    #endregion
                }
                return Tuple.Create(itemTable, documentTypeLanguage, contacts);
            }
        }
        private static Dictionary<string, string> Get_contactKeysValues(List<Word> words, double col3_qtyLeft, byte margin)
        {
            var keysAndValues = new Dictionary<string, string>();
            words.Sort((w1, w2) =>
            {
                int lvl1 = w2.BoundingBox.Bottom.CompareTo(w1.BoundingBox.Bottom);
                if (lvl1 != 0) return lvl1;
                int lvl2 = w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left);
                return lvl2;
            });
            double lineHeight = words.Average(w => w.BoundingBox.Height);
            var words_byLine = new Dictionary<byte, List<Word>>();
            foreach (var word in words)
            {
                double Y = word.BoundingBox.Bottom;
                byte lineNbr = Convert.ToByte(Y / lineHeight);
                if (!words_byLine.ContainsKey(lineNbr)) words_byLine[lineNbr] = new List<Word>();
                words_byLine[lineNbr].Add(word);
            }
            words_byLine = words_byLine.OrderByDescending(y => y.Key).ToDictionary(k => k.Key, y => y.Value);
            foreach (var lineNbr in words_byLine.Keys)
            {
                words_byLine[lineNbr].Sort((w1, w2) => { return w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left); });
                var firstWord = words_byLine[lineNbr].First();
                var words_thisLine = Words_inLine(firstWord.BoundingBox, words, margin);
                var words_AboveBelow = new List<Word>();
                foreach (var isAbove in words_thisLine.Keys)
                    words_AboveBelow.AddRange(words_thisLine[isAbove]);
                var leftSide = Get_bySide(words_AboveBelow, col3_qtyLeft);
                foreach (var newKeyValue in leftSide)
                    if (!keysAndValues.ContainsKey(newKeyValue.Key)) keysAndValues.Add(newKeyValue.Key, newKeyValue.Value);
                var rightSide = Get_bySide(words_AboveBelow, col3_qtyLeft, false);
                foreach (var newKeyValue in rightSide)
                    if (!keysAndValues.ContainsKey(newKeyValue.Key)) keysAndValues.Add(newKeyValue.Key, newKeyValue.Value);
            }
            return keysAndValues;
        }
        private static Dictionary<string, string> Get_bySide(List<Word> words, double col3_qtyLeft, bool isLeftSide = true)
        {
            var keysAndValues = new Dictionary<string, string>();
            var words_OneSide = new List<Word>();
            if (isLeftSide)
                words_OneSide.AddRange(words.Where(w => w.BoundingBox.Right < col3_qtyLeft));
            else
                words_OneSide.AddRange(words.Where(w => w.BoundingBox.Right > col3_qtyLeft));
            var words_colon = new List<Word>(words_OneSide.Where(w => w.Text.Contains(":")));
            if (words_colon.Any())
            {
                var rect_colon = words_colon[0].BoundingBox;
                var words_colonLeft = new List<Word>(words_OneSide.Where(w => w.BoundingBox.Left <= rect_colon.Left));
                var words_colonRight = new List<Word>(words_OneSide.Where(w => w.BoundingBox.Left > rect_colon.Left));
                if (words_colonRight.Any())
                {
                    string wordKey = string.Join(" ", words_colonLeft.Select(w => w.Text));
                    string wordValue = string.Join(" ", words_colonRight.Select(w => w.Text));
                    if (!keysAndValues.ContainsKey(wordKey)) keysAndValues[wordKey] = wordValue;
                }
            }
            else
            {
                /// could be no words on the left needed
                /// ...OR the wrapped text
                /// SOUMISSION #:
                ///     NSOPP6009                           <-- no colon
                /// address lines ( city, province, postal code) which are NOT delimited with a colon
                /// Adresse: 3530 RUE JEAN TALON OUEST
                /// MONTREAL, QC H3R 2G3                    <-- no colon

                if (words_OneSide.Any())
                {
                    string wrappedText = string.Join(" ", words_OneSide).Trim();
                    if (Regex.IsMatch(wrappedText, "[ABCEGHJ-NPRSTVXY]\\d[ABCEGHJ-NPRSTV-Z][ -]?\\d[ABCEGHJ-NPRSTV-Z]\\d"))
                        keysAndValues.Add("CityProvincePostal", wrappedText);
                    if (Regex.IsMatch(wrappedText, "[A-Z]{3}PP[0-9]{4}"))
                        keysAndValues.Add("quote#", wrappedText);
                }
            }
            return keysAndValues;
        }
        internal void Send_gmail()
        {
            // https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4POmWnx38P9p4UgNgPHjEGTYiuFPrxoOX9MSGslj7mZVhJ6k3-pvQUxFYVQHojrNkiQx0t9YosWMXcbbUxfGg5_bg1PFA
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("seanglover.spg@gmail.com");
                mail.To.Add("seanglover.spg@gmail.com");
                mail.Subject = "Hello World";
                mail.Body = "<h1>Hello</h1>";
                mail.IsBodyHtml = true;
                //mail.Attachments.Add(new Attachment("C:\\file.zip"));

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("seanglover.spg@gmail.com", "db2sql01+Luvindam0nkey");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }
        }
        private static string Printout()
        {
            string client = null;
            string order = null;
            string contact = null;
            string phone = null;
            string repName = null;
            string repEmail = null;

            const byte halfway = 50; // 50 is halfway mark
            const byte indent = 8; // 50 is halfway mark
            string rightPad = new string(' ', halfway);
            string leftPad = new string(' ', indent);
            var all = new List<string>();
            var line1a = $"Client: {client?.Trim()}";
            var line1b = $"Order: {order?.Trim()}";
            var line1 = (line1a + rightPad).Substring(0, halfway) + (line1b + rightPad).Substring(0, halfway);
            all.Add((leftPad + line1).Substring(0, halfway * 2));

            var line2a = $"Contact: {contact?.Trim()}";
            var line2b = $"Phone: {phone?.Trim()}";
            var line2 = (line2a + rightPad).Substring(0, halfway) + (line2b + rightPad).Substring(0, halfway);
            all.Add((leftPad + line2).Substring(0, halfway * 2));

            var line3a = $"Rep: {repName?.Trim()}";
            var line3b = $"email: {repEmail?.Trim()}";
            var line3 = (line3a + rightPad).Substring(0, halfway) + (line3b + rightPad).Substring(0, halfway);
            all.Add((leftPad + line3).Substring(0, halfway * 2));

            const byte col1 = 12;
            const byte col2 = 81;
            const byte col3 = 3;
            all.Add("┏" + new string('━', col1) + "┳" + new string('━', col2) + "┳" + new string('━', col3) + "┓");
            all.Add("┃" + ("Item#" + new string(' ', col1)).Substring(0, col1) + "┃" + ("Desc." + new string(' ', col2)).Substring(0, col2) + "┃Qty┃");
            all.Add("┣" + new string('━', col1) + "╋" + new string('━', col2) + "╋" + new string('━', col3) + "┫");
            //foreach (var tableRow in tableRows)
            //{
            //    var rowDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(tableRow.Value);
            //    string cell1_item = rowDict[ColumnNames[0]];
            //    string cell2_desc = rowDict[ColumnNames[1]];
            //    string cell3_qty = rowDict.ContainsKey(ColumnNames[2]) ? rowDict[ColumnNames[2]] : string.Empty; // may not contain (ex. MILKY WAY CARPET KIT CONSISTS OF)
            //    all.Add("┃" + (cell1_item + new string(' ', col1)).Substring(0, col1) + "┃" + (cell2_desc + new string(' ', col2)).Substring(0, col2) + $"┃" + (cell3_qty + new string(' ', col3)).Substring(0, col3) + "┃");
            //}
            all.Add("┗" + new string('━', col1) + "┻" + new string('━', col2) + "┻" + new string('━', col3) + "┛");
            return string.Join(Environment.NewLine, all);
        }
        private struct Item
        {
            public string Code { get; set; }
            public string Description { get; set; }
            public double Qty { get; set; }
            public bool Complete { get { return !(Code == null | Description == null | Qty == 0); } }
            public override string ToString()
            {
                string code = Code ?? "null";
                string codeString = (code + new string(' ', 12)).Substring(0, 12);
                return $"[{codeString}]*[{Qty:000}] {Description}";
            }
        }
    }
}