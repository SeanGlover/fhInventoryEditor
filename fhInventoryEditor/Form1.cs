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
using System.Drawing;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Geometry;
using DataTableAsync;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Net.Mail;
using System.Net;
using System.Xml.Linq;

namespace fhInventoryEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static readonly DirectoryInfo jobsFolder = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\");
        private static readonly string accountName = "Centre Beaubien"; // CentreLeCap, Centre Beaubien
        private static readonly DirectoryInfo jobInfo = new DirectoryInfo($"{jobsFolder.FullName}{accountName}\\");
        private static readonly string htmlPath = $"{jobInfo.FullName}a_jobSummary\\index.html";
        private readonly HtmlDocument htmlEditor = new HtmlDocument();

        public enum PageRegion { none, contact, table_disclaimer, table_hdr, table_data, footer }
        internal bool FormIsFrench { get; private set; }
        internal bool FormIsDelivery { get; private set; }
        internal List<string> ColumnNames { get; private set; }
        private readonly Table contactTable = new Table();
        private readonly Table itemTable = new Table();
        private readonly Dictionary<PageRegion, Dictionary<byte, string>> regions = new Dictionary<PageRegion, Dictionary<byte, string>>
        {
            [PageRegion.contact] = new Dictionary<byte, string>(),
            [PageRegion.table_hdr] = new Dictionary<byte, string> { [0] = string.Empty },
            [PageRegion.table_data] = new Dictionary<byte, string>()
        };

#region" contact info "
internal string Client
        {
            get { return client; }
            private set { } }
        private string client;
        internal string Order
        {
            get { return order; }
            private set { }
        }
        private string order;
        internal string Contact
        {
            get { return contact; }
            private set { }
        }
        private string contact;
        internal string Phone
        {
            get { return phone; }
            private set { }
        }
        private string phone;
        internal string RepName
        {
            get { return repName; }
            private set { }
        }
        private string repName;
        internal string RepEmail
        {
            get { return repEmail; }
            private set { }
        }
        private string repEmail;
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            var item = itemTable.Columns.Add("item", typeof(string));
            var desc = itemTable.Columns.Add("desc", typeof(string));
            var qty = itemTable.Columns.Add("qty", typeof(double));
            itemTable.PrimaryKeys = new Table.Column[] { item };

            var leftKey = contactTable.Columns.Add("leftKey", typeof(string));
            var leftValue = contactTable.Columns.Add("leftValue", typeof(string));
            var rightKey = contactTable.Columns.Add("rightKey", typeof(string));
            var rightValue = contactTable.Columns.Add("rightValue", typeof(string));
            contactTable.PrimaryKeys = new Table.Column[] { leftKey };

            List<FileInfo> allFiles = new List<FileInfo>(jobsFolder.EnumerateFiles("*.pdf", SearchOption.AllDirectories));
            List<FileInfo> allDeliveries = new List<FileInfo>(jobsFolder.EnumerateFiles("*.pdf", SearchOption.AllDirectories).Where(f => Regex.IsMatch(f.Name.ToLowerInvariant(), "delivery|livraison")));
            allDeliveries.Sort((f1, f2) => { return string.Compare(f1.Name, f2.Name); });
            List<FileInfo> jobDelivery = new List<FileInfo>(jobInfo.EnumerateFiles("*.pdf").Where(f => Regex.IsMatch(f.Name.ToLowerInvariant(), "delivery|livraison")));
            #region" delivery examples "
            const bool byIndex = true;
            const byte deliveryIndex = 1;
            const string cadens = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/Cadens%20Lighthouse/NS0PP5662%20Caden's%20Lighthouse%20-%20Delivery%20Verification%20Form.pdf";
            const string leCap = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/CentreLeCap/APOPP6975%20Delivery%20Checklist.pdf";
            const string jLeger = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/JulesLeger/CENTRE%20JULES%20LEGER%20-%20DELIVERY%20CHECKLIST%20-%20PROJECT%20NSOPP6557.pdf";
            const string bbienA = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/Centre%20Beaubien/NSOPP6331%20FORMULAIRE%20DE%20V%C3%89RIFICATION%20DE%20LIVRAISON.pdf";
            const string evasion = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/Centre%20Evasion/CENTRE%20%C3%89VASION_delivery.pdf";
            const string yQuote = "file:///C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/Yaldei/NSOPP5357%20-%20YALDEI%20CENTER%20-%20SNOEZELEN%20MULTI-SENSORY%20ROOM%20-%20QUOTE%20P0810161%20revised.pdf";
            string tabURL = System.Web.HttpUtility.UrlDecode(yQuote);
            FileInfo deliveryForm = byIndex ? allDeliveries[deliveryIndex] : new FileInfo(tabURL.Replace("file:///", string.Empty));
            /// 0=cadens            NO "consists of:" wrapping... might be better to go for each line
            /// 1=cdbc restigouche  100%
            /// 2=beaubien (a)      100%
            /// 3=beaubien (b)      100%
            /// 4=leCap             NO wrap
            /// 5=julesLeger        NO wrap
            /// 6=onyva-hawkesbury  100%
            /// 7=onyva-rockland    100%
            /// 8=angelica          100%
            /// 9=yaldei            
            #endregion
            foreach (FileInfo delivery in allFiles.Skip(0).Take(1000))
            {
                Parse_deliveryForm(delivery);
                //Debugger.Break();
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
        private static string CleanDescription(string descriptionIn)
        {
            string pad = new string(' ', 8);
            descriptionIn += pad;
            string newDescription = descriptionIn;
            newDescription = Regex.Replace(newDescription, "\\sTRK\\s", "     ") + pad; // dont know what these are
            newDescription = Regex.Replace(newDescription, " N[A-Z][0-9]{2} ", pad) + pad; // NS15 or NF13 etc
            while (Regex.IsMatch(newDescription, " ([A-Z])\\1 "))
            {
                newDescription = Regex.Replace(newDescription, " ([A-Z])\\1 ", pad) + pad; // AA, BB, etc
            }
            newDescription = Regex.Replace(newDescription, " {2,}", " ") + pad;
            newDescription = newDescription.Trim();
            return newDescription;
        }
        private static Dictionary<bool, List<Word>> Words_inLine(PdfRectangle rectIn, List<Word> words, byte margin = 0)
        {
            var xxx = new PdfRectangle(new PdfPoint(0, rectIn.BottomLeft.Y - margin), new PdfPoint(1000, rectIn.TopRight.Y + margin));
            var wordsInRects = new List<Word>(words.Where(w => w.BoundingBox.IntersectsWith(xxx)));
            var words_inLine =
(from ltr in wordsInRects
 let isAbove = ltr.BoundingBox.Top > xxx.Centroid.Y
 group ltr by isAbove into lineGroup
 orderby lineGroup.Key descending
 select new
 {
     above = lineGroup.Key,
     words = new List<Word>(lineGroup.OrderBy(w => w.BoundingBox.Left))
 }).ToDictionary(k => k.above, v => v.words);
            return words_inLine;
        }

        private void Parse_deliveryForm(FileInfo jobinfo, bool openFile = false)
        {
            if (jobinfo != null)
            {
                if (openFile) Process.Start($"{jobinfo.FullName}");

                const byte pageWidth = 112;
                string emptyLine = new string(' ', pageWidth);
                regions[PageRegion.table_data].Clear();
                itemTable.Rows.Clear();
                contactTable.Rows.Clear();
                try
                {
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
                            List<Word> tableWords = new List<Word>();

                            byte pageNbr = Convert.ToByte(page.Number);
                            if (pageNbr == 1)
                            {
                                FormIsFrench = pageText.ToLowerInvariant().Contains("vérification");
                                FormIsDelivery = FormIsFrench ? pageText.ToLowerInvariant().Contains("livraison") : pageText.ToLowerInvariant().Contains("delivery");
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
                                    if (r1.Height > 0 & r1.Height < 2) pdfRects_linesHorizontal.Add(r1);
                                    if (r1.Width > 0 & r1.Width < 2)
                                    {
                                        pdfRects_linesVertical.Add(r1);
                                        if (thinnestVerticalLine > r1.Width) thinnestVerticalLine = r1.Width;
                                    }
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
                                var table_linesVertical =
                (from vLine in pdfRects_linesVertical
                 where Math.Round(vLine.Width, 2) == thinnestVerticalLine
                 let lineLeft = vLine.Left
                 group vLine by lineLeft into leftGroup
                 orderby leftGroup.Key ascending
                 select new
                 {
                     LineLeft = leftGroup.Key,
                     Lines = new List<PdfRectangle>(leftGroup.OrderBy(x => x.Height))
                 }
                 ).ToDictionary(k => k.LineLeft, v => v.Lines);
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
                                if (FormIsDelivery)
                                {
                                    double VERIFICATION = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "v(e|é)rification", RegexOptions.IgnoreCase))).FirstOrDefault().BoundingBox.Bottom;
                                    double PLEASENOTE = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "NOTE(R){0,1}:"))).FirstOrDefault().BoundingBox.Top;
                                    var wordsAboveDisclaimer = new List<Word>(wordsAboveTable.Where(w => w.BoundingBox.Bottom > PLEASENOTE));
                                    var wordsContact = new List<Word>(wordsAboveDisclaimer.Where(w => w.BoundingBox.Top < VERIFICATION));
                                    var wordsContactKeys = new List<Word>(wordsContact.Where(w => w.Text.EndsWith(":")));
                                    var wordsContactKeysLeft = new List<Word>(wordsContactKeys.Where(w => w.BoundingBox.Right < col3_qtyLeft));
                                    foreach (var contactKey in wordsContactKeysLeft)
                                    {
                                        var words_thisLine = Words_inLine(contactKey.BoundingBox, words, 5);
                                        var words_above = new List<Word>();
                                        if (words_thisLine.ContainsKey(true)) words_above.AddRange(words_thisLine[true].Where(w => w.BoundingBox.Left > contactKey.BoundingBox.Right & w.BoundingBox.Right < col3_qtyLeft));
                                        var words_below = new List<Word>();
                                        if (words_thisLine.ContainsKey(false)) words_below.AddRange(words_thisLine[false].Where(w => w.BoundingBox.Left > contactKey.BoundingBox.Right & w.BoundingBox.Right < col3_qtyLeft));
                                        var words_thisLeft = new List<Word>(words_above.Union(words_below));
                                        if (words_thisLeft.Any())
                                        {
                                            string wordsLeftText = string.Join(" ", words_thisLeft.Select(w => w.Text)).Trim();
                                            if (Regex.IsMatch(contactKey.Text, "(customer|client):", RegexOptions.IgnoreCase)) client = wordsLeftText;
                                            else if (Regex.IsMatch(contactKey.Text, "contact:", RegexOptions.IgnoreCase)) contact = wordsLeftText;
                                            else if (Regex.IsMatch(contactKey.Text, "rep:", RegexOptions.IgnoreCase)) repName = wordsLeftText;
                                        }
                                    }
                                    var wordsContactKeysRight = new List<Word>(wordsContactKeys.Where(w => w.BoundingBox.Left > col3_qtyLeft));
                                    foreach (var contactKey in wordsContactKeysRight)
                                    {
                                        var words_thisLine = Words_inLine(contactKey.BoundingBox, words, 5);
                                        var words_above = new List<Word>();
                                        if (words_thisLine.ContainsKey(true)) words_above.AddRange(words_thisLine[true].Where(w => w.BoundingBox.Left > contactKey.BoundingBox.Right));
                                        var words_below = new List<Word>();
                                        if (words_thisLine.ContainsKey(false)) words_below.AddRange(words_thisLine[false].Where(w => w.BoundingBox.Left > contactKey.BoundingBox.Right));
                                        var words_thisRight = new List<Word>(words_above.Union(words_below));
                                        if (words_thisRight.Any())
                                        {
                                            string wordsRightText = string.Join(" ", words_thisRight.Select(w => w.Text)).Trim();
                                            if (Regex.IsMatch(contactKey.Text, "(order|commande):", RegexOptions.IgnoreCase)) order = wordsRightText;
                                            else if (Regex.IsMatch(contactKey.Text, "phone:", RegexOptions.IgnoreCase)) phone = wordsRightText;
                                            else if (Regex.IsMatch(contactKey.Text, "e-mail:", RegexOptions.IgnoreCase)) repEmail = wordsRightText;
                                        }
                                    }
                                }
                                else
                                {
                                    //Debugger.Break();
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
                                    regions[PageRegion.table_hdr][0] = string.Join(" ", colNames.Keys);
                                    tableWords.AddRange(words.Where(w => w.BoundingBox.Top < itemNbr.BoundingBox.Bottom));
                                }
                                #endregion
                            }
                            else
                            {
                                pgRegion = PageRegion.table_data; // page2+ is a continuation of the table from page1 etc}
                                tableWords.AddRange(words);
                                //Debugger.Break();
                            }

                            #region" fonts --> size(w,h) "
                            Dictionary<string, List<Word>> fonts =
        (from w in words
         let fontName = w.FontName
         group w by fontName into fontGroup
         orderby fontGroup.Key
         select new { FontName = fontGroup.Key, Words = new List<Word>(fontGroup) }).ToDictionary(k => k.FontName, v => v.Words);
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
                            var letters_thisPage =
            (from ltr in page.Letters
             let lineNbr = Convert.ToByte(ltr.StartBaseLine.Y / fontHeight)
             group ltr by lineNbr into lineGroup
             orderby lineGroup.Key descending // order the dictionary by lineNbr DESC (Max(Location.Y)=> page top, Min(Location.Y)=> page bottom)
             select new
             {
                 LineNumber = lineGroup.Key,
                 Lines = (from ln in lineGroup
                          let Y = ln.StartBaseLine.Y
                          group ln by Y into yGrp
                          orderby yGrp.Min(yg => yg.StartBaseLine.X)
                          let Xs = new List<double>(yGrp.Select(x => x.StartBaseLine.X).OrderBy(x => x))
                          select new
                          {
                              yPos = yGrp.Key,
                              Letters = yGrp.OrderBy(x => x.StartBaseLine.X).ToDictionary(k => Xs.IndexOf(k.StartBaseLine.X), v => v)
                          }).ToDictionary(k => k.yPos, v => v.Letters)
             }
             ).ToDictionary(k => k.LineNumber, v => v.Lines);
                            letters_byPage[pageNbr] = letters_thisPage;
                            #endregion

                            var lines_thisPage = new Dictionary<double, string>();
                            var lines = new List<string>();
                            var tableRows = regions[PageRegion.table_data];
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
                                        var columnWords = words_byColumn.ToDictionary(k => k.Key, v => CleanDescription(string.Join(" ", v.Value.Select(w => w.Text))), StringComparer.OrdinalIgnoreCase);
                                        tableRows.Add((byte)tableRows.Count, JsonConvert.SerializeObject(columnWords, Formatting.None));
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
                                    if (!line.Any()) isTableEnd = true;
                                    if (isTableEnd)
                                        pgRegion = PageRegion.footer;
                                }
                            }
                            //Debugger.Break();
                        }

                        #region" save .txt file "
                        const byte halfway= 50; // 50 is halfway mark
                        const byte indent = 8; // 50 is halfway mark
                        string rightPad = new string(' ', halfway);
                        string leftPad = new string(' ', indent);
                        var all = new List<string>();
                        var line1a = $"Client: {client?.Trim()}";
                        var line1b = $"Order: {order?.Trim()}";
                        var line1 = (line1a + rightPad).Substring(0, halfway) + (line1b + rightPad).Substring(0, halfway);
                        all.Add((leftPad + line1).Substring(0, halfway * 2));
                        contactTable.Rows.Add(new string[] { "Client:", client, "Order:", order });

                        var line2a = $"Contact: {contact?.Trim()}";
                        var line2b = $"Phone: {phone?.Trim()}";
                        var line2 = (line2a + rightPad).Substring(0, halfway) + (line2b + rightPad).Substring(0, halfway);
                        all.Add((leftPad + line2).Substring(0, halfway * 2));
                        contactTable.Rows.Add(new string[] { "Contact:", contact, "Phone:", phone });

                        var line3a = $"Rep: {repName?.Trim()}";
                        var line3b = $"email: {repEmail?.Trim()}";
                        var line3 = (line3a + rightPad).Substring(0, halfway) + (line3b + rightPad).Substring(0, halfway);
                        all.Add((leftPad + line3).Substring(0, halfway * 2));
                        contactTable.Rows.Add(new string[] { "Rep name:", repName, "Rep email:", repEmail });

                        const byte col1 = 12;
                        const byte col2 = 81;
                        const byte col3 = 3;
                        all.Add("┏" + new string('━', col1) + "┳" + new string('━', col2) + "┳" + new string('━', col3) + "┓");
                        all.Add("┃" + ("Item#" + new string(' ', col1)).Substring(0, col1) + "┃" + ("Desc." + new string(' ', col2)).Substring(0, col2) + "┃Qty┃");
                        all.Add("┣" + new string('━', col1) + "╋" + new string('━', col2) + "╋" + new string('━', col3) + "┫");
                        foreach (var item in regions[PageRegion.table_data])
                        {
                            var rowDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(item.Value);
                            string cell1_item = rowDict[ColumnNames[0]];
                            string cell2_desc = rowDict[ColumnNames[1]];
                            string cell3_qty = rowDict.ContainsKey(ColumnNames[2]) ? rowDict[ColumnNames[2]] : string.Empty; // may not contain (ex. MILKY WAY CARPET KIT CONSISTS OF)
                            all.Add("┃" + (cell1_item + new string(' ', col1)).Substring(0, col1) + "┃" + (cell2_desc + new string(' ', col2)).Substring(0, col2) + $"┃" + (cell3_qty + new string(' ', col3)).Substring(0, col3) + "┃");
                            itemTable.Rows.Add(new object[] { cell1_item, cell2_desc, cell3_qty });
                        }
                        all.Add("┗" + new string('━', col1) + "┻" + new string('━', col2) + "┻" + new string('━', col3) + "┛");
                        string printout = string.Join(Environment.NewLine, all);
                        string htmlContact = contactTable.HTML;
                        string htmlData = itemTable.HTML;
                        string html = string.Join(Environment.NewLine, new string[] { htmlContact, htmlData });
                        itemTable.Name = string.Join("■", new string[] { client, order, contact, phone, repName, repEmail });
                        string legalOrder = Regex.Replace(order ?? string.Empty, "[\\\\/:\"*?<>|]+", "^");
                        string newFilePath = $"{jobinfo.Directory.FullName}\\{jobinfo.Directory.Name}_deliveryList[{legalOrder}].txt";
                        File.WriteAllText(newFilePath, JsonConvert.SerializeObject(itemTable, Formatting.Indented));
                        File.Move(jobinfo.FullName, newFilePath.Replace(".txt", ".pdf"));
                        #endregion
                    }
                }
                catch { }
            }
        }
        private void Open_quotes(FileInfo jobinfo)
        {
            using (PdfDocument document = PdfDocument.Open(jobinfo.FullName))
            {
                foreach (Page page in document.GetPages())
                {
                    string pageText = page.Text;
                    if (Regex.IsMatch(pageText, "(QUOTE|SOUMISSION) {0,}#:", RegexOptions.IgnoreCase))
                    {
                        Process.Start($"{jobinfo.FullName}");
                        break;
                    }
                }
            }
        }
        private void Send_gmail()
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