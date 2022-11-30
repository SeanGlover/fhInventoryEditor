using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.IO;
//using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Page = UglyToad.PdfPig.Content.Page;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace fhInventoryEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static readonly DirectoryInfo jobsFolder = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\");
        private static readonly string accountName = "CentreLeCap"; // CentreLeCap, Centre Beaubien
        private static readonly DirectoryInfo jobInfo = new DirectoryInfo($"{jobsFolder.FullName}{accountName}\\");
        private static readonly string htmlPath = $"{jobInfo.FullName}a_jobSummary\\index.html";
        private readonly HtmlDocument htmlEditor = new HtmlDocument();

        public enum PageRegion { none, contact, table_disclaimer, table_hdr, table_data, footer }
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
        private Dictionary<byte, Item> items { get; } = new Dictionary<byte, Item>();

        private void Form1_Load(object sender, EventArgs e)
        {
            htmlEditor.LoadHtml(File.ReadAllText(htmlPath));

            //htmlEditor.GetElementbyId("businessName").SetAttributeValue("value", "Centre Le Cap");
            //htmlEditor.GetElementbyId("purchaseOrder").SetAttributeValue("value", "");

            //using (StreamWriter sw = new StreamWriter(htmlPath))
            //{
            //    htmlEditor.Save(sw);
            //}
            Parse_deliveryForm();
            Debugger.Break();
        }
        private void Parse_deliveryForm()
        {
            List<FileInfo> deliveryInfo = new List<FileInfo>(jobInfo.EnumerateFiles("*.pdf").Where(f => Regex.IsMatch(f.Name.ToLowerInvariant(), "delivery|livraison")));
            if (deliveryInfo.Any())
            {
                FileInfo deliveryForm = deliveryInfo.First();
                const byte pageWidth = 112;
                string emptyLine = new string(' ', pageWidth);

                using (PdfDocument document = PdfDocument.Open(deliveryForm.FullName))
                {
                    var fonts_byPage = new Dictionary<byte, Dictionary<string, List<Word>>>();
                    var lines_byPage = new Dictionary<byte, Dictionary<byte, string>>(); // {pgNbr, {lineNbr, Line}}
                    var words_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<double, List<Word>>>>();
                    var letters_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<double, Dictionary<int, Letter>>>>();
                    var consecutiveLetters_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<byte, Dictionary<int, Letter>>>>();
                    var regions_byPage = new Dictionary<byte, Dictionary<PageRegion, Dictionary<byte, string>>>();
                    var items_byPage = new Dictionary<byte, Dictionary<byte, Item>>();

                    foreach (Page page in document.GetPages())
                    {
                        var boxes = page.ExperimentalAccess.Paths.Select(p => p.GetBoundingRectangle()).Where(bb => bb.HasValue).Select(bb => bb.Value).ToList();

                        byte pageNbr = Convert.ToByte(page.Number);
                        PageRegion pgRegion = new PageRegion();
                        List<Word> words = new List<Word>(page.GetWords());
                        lines_byPage[pageNbr] = new Dictionary<byte, string>();
                        consecutiveLetters_byPage[pageNbr] = new Dictionary<byte, Dictionary<byte, Dictionary<int, Letter>>>();
                        regions_byPage[pageNbr] = new Dictionary<PageRegion, Dictionary<byte, string>>();

                        #region" fonts --> size(w,h) "
                        Dictionary<string, List<Word>> fonts =
    (from w in words
     let fontName = w.FontName
     group w by fontName into fontGroup
     orderby fontGroup.Key
     select new { FontName = fontGroup.Key, Words = new List<Word>(fontGroup) }).ToDictionary(k => k.FontName, v => v.Words);
                        fonts_byPage[pageNbr] = fonts;
                        Dictionary<string, double> spacing_byFont = new Dictionary<string, double>();
                        foreach (string fontName in fonts.Keys)
                        {
                            double fw = 0;
                            List<Word> fontWords = fonts[fontName];
                            fontWords.Sort((y1, y2) => y2.Letters.Count.CompareTo(y1.Letters.Count));
                            if (fontWords.Any())
                            {
                                List<Letter> longestWordLetters = new List<Letter>(fontWords.First().Letters);
                                List<double> fontXs = new List<double>(longestWordLetters.Select(x => Math.Round(x.Location.X, 1)).Distinct());
                                fw = Math.Round((fontXs.Max() - fontXs.Min()) / (fontXs.Count - 1), 1);
                            }
                            spacing_byFont.Add(fontName, fw);
                        }
                        fonts = fonts.OrderByDescending(f => f.Value.Count).ToDictionary(k => k.Key, v => v.Value);
                        string mostCommonFont = fonts.Keys.FirstOrDefault().ToString();
                        double fontWidth = spacing_byFont[mostCommonFont];
                        const byte fontHeight = 9;
                        #endregion
                        #region" words_thisPage "
                        var words_thisPage =
    (from wrd in words
     let lineNbr = Convert.ToByte(wrd.BoundingBox.Bottom / fontHeight)
     group wrd by lineNbr into lineGroup
     orderby lineGroup.Key descending // order the dictionary by lineNbr DESC (Max(Location.Y)=> page top, Min(Location.Y)=> page bottom)
     select new
     {
         LineNumber = lineGroup.Key,
         Lines = (from ln in lineGroup
                  let Y = ln.BoundingBox.Bottom
                  group ln by Y into yGrp
                  orderby yGrp.Min(yg => yg.BoundingBox.Left)
                  select new { yPos = yGrp.Key, Words = new List<Word>(yGrp.OrderBy(y => y.BoundingBox.Left)) }).ToDictionary(k => k.yPos, v => v.Words)
     }
    ).ToDictionary(k => k.LineNumber, v => v.Lines);
                        words_byPage[pageNbr] = words_thisPage;
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

                        Dictionary<double, string> lines_thisPage = new Dictionary<double, string>();
                        List<string> lines = new List<string>();
                        string wordDelimiter = "\t";

                        foreach (byte lineNbr in letters_thisPage.Keys)
                        {
                            byte groupIndex = 0;
                            Dictionary<byte, Dictionary<int, Letter>> consecutiveLetters = new Dictionary<byte, Dictionary<int, Letter>>();
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
                            double left = consecutiveLetters[0][0].StartBaseLine.X;
                            string lineData = $"{lineNbr:000}_{left:000.0}|" + string.Join("■", consecutiveLetters.Select(cl => string.Join(string.Empty, cl.Value.Select(l => l.Value.Value))));
                            lines.Add(lineData);
                            
                            List<string> wordGroups = new List<string>();
                            foreach (var lettersGroup in consecutiveLetters)
                            {
                                /// calculate the relative position of each lettergroup (word)
                                /// ex. page width = 600px, 1st word starts at 150px (25%)
                                /// 112 characters wide would be 28th character
                                string word = string.Join(string.Empty, lettersGroup.Value.OrderBy(l => l.Key).Select(l => l.Value.Value));
                                if (pgRegion == PageRegion.table_hdr | pgRegion == PageRegion.table_data)
                                    word = $"●{consecutiveLetters[lettersGroup.Key][0].StartBaseLine.X:N1}● {word}";
                                wordGroups.Add(word);
                            }
                            
                            string line = string.Join(wordDelimiter, wordGroups);
                            lines_thisPage[lineNbr] = line;
                            lines_byPage[pageNbr].Add((byte)lines_byPage[pageNbr].Count, line);

                            if (lineData.Contains("ITEM #"))
                            {
                                // ITEM # | DESCRIPTION | QUANTITY | RECEIVED | MISSING
                                // ITEM # | DESCRIPTION | QUANTITÉ | REÇU | MANQUANT
                                pgRegion = PageRegion.table_hdr;
                                regions_byPage[pageNbr][pgRegion] = new Dictionary<byte, string>();
                                regions_byPage[pageNbr][pgRegion][0] = line; // table header is 1 line only
                                wordDelimiter = "■";
                            }
                            else if (pgRegion == PageRegion.none & Regex.IsMatch(line, "v(e|é)rification", RegexOptions.IgnoreCase))
                                pgRegion = PageRegion.contact;

                            else if (pgRegion == PageRegion.contact)
                            {
                                var contactRegion = regions_byPage[pageNbr][pgRegion];
                                if (Regex.IsMatch(line, "(customer|client):", RegexOptions.IgnoreCase))
                                {
                                    string[] client_Order = Regex.Split(line.Trim(), "(Customer|Client|Order|Commande):", RegexOptions.IgnoreCase);
                                    client = client_Order[2].Trim();
                                    order = client_Order[4].Trim();
                                    contactRegion.Add(0, JsonConvert.SerializeObject(new Dictionary<string, string>() { { "client", client }, { "order#", order } }));
                                }
                                else if (Regex.IsMatch(line, "Contact:"))
                                {
                                    string[] contact_Phone = Regex.Split(line.Trim(), "(téléphone|phone):", RegexOptions.IgnoreCase);
                                    contact = Regex.Match(contact_Phone[0], "(?<=contact:).*", RegexOptions.IgnoreCase).Value.Trim();
                                    phone = Regex.Match(line, "(?<=phone:).*", RegexOptions.IgnoreCase).Value.Trim();
                                    contactRegion.Add(1, JsonConvert.SerializeObject(new Dictionary<string, string>() { { "contact", contact }, { "phone#", phone } }));
                                }
                                else if (line.Contains("E-Mail:"))
                                {
                                    string[] rep_Email = Regex.Split(line.Trim(), "E-Mail:", RegexOptions.IgnoreCase);
                                    repName = Regex.Replace(rep_Email[0], "flaghouse", string.Empty, RegexOptions.IgnoreCase).Trim();
                                    repEmail = rep_Email[1].Trim();
                                    contactRegion.Add(2, JsonConvert.SerializeObject(new Dictionary<string, string>() { { "rep", repName }, { "email", repEmail } }));
                                }
                            }
                            else if (pgRegion == PageRegion.table_hdr)
                                pgRegion = PageRegion.table_data;

                            if (!regions_byPage[pageNbr].ContainsKey(pgRegion)) regions_byPage[pageNbr][pgRegion] = new Dictionary<byte, string>();

                            if (pgRegion == PageRegion.table_data)
                            {
                                // [117.5] SIGNATURE ■[336.8] DATE
                                bool isTableEnd = line.Contains("SIGNATURE") & line.Contains("DATE");
                                if (line.Contains("NOM IMPRIMÉ") | line.Contains("PRINTED NAME")) isTableEnd = true;
                                if (line.Contains("VEUILLEZ ENVOYER")) isTableEnd = true;
                                if (!line.Any()) isTableEnd = true;

                                if (isTableEnd)
                                    pgRegion = PageRegion.footer;

                                else
                                {
                                    var tableRegion = regions_byPage[pageNbr][pgRegion];
                                    tableRegion.Add(lineNbr, line);
                                }
                            }
                        }

                        var tableRows = regions_byPage[pageNbr][PageRegion.table_data];
                        var wordLefts = new Dictionary<double, List<Tuple<byte, string>>>();
                        foreach (var tableRow in tableRows)
                        {
                            var xs = Regex.Matches(tableRow.Value, "●[0-9.]{1,}●{1,}");
                            foreach (Match x in xs)
                            {
                                string nbr = x.Value.Replace("●", string.Empty);
                                double wordLeft = double.Parse(nbr);
                                if (!wordLefts.ContainsKey(wordLeft)) wordLefts[wordLeft] = new List<Tuple<byte, string>>();
                                string[] splitRow = tableRow.Value.Split(new string[] { x.Value, "■" }, StringSplitOptions.RemoveEmptyEntries);
                                wordLefts[wordLeft].AddRange(splitRow.Where(sr => !sr.Contains("●")).Select(sr => Tuple.Create(tableRow.Key, sr.Trim())));
                            }
                        }
                        wordLefts = wordLefts.OrderByDescending(wl => wl.Key).ToDictionary(k => k.Key, v => v.Value);
                        var columns = new Dictionary<string, List<Tuple<byte, string>>>();
                        var col3_x = wordLefts.First().Key;
                        var col2_x = wordLefts.Skip(1).First().Key;
                        foreach (var row in wordLefts)
                        {
                            string colName = row.Key == col3_x ? "3_Qty" : row.Key == col2_x ? "2_Description" : "1_Code";
                            if (!columns.ContainsKey(colName)) columns.Add(colName, new List<Tuple<byte, string>>());
                            columns[colName].AddRange(row.Value);
                        }
                        var lineNbrs = new List<byte>();
                        foreach (var col in columns)
                        {
                            col.Value.Sort((v1, v2) => v2.Item1.CompareTo(v1.Item1));
                            lineNbrs.AddRange(col.Value.Select(c => c.Item1));
                        }
                        lineNbrs = lineNbrs.Distinct().ToList();
                        lineNbrs.Sort((l1, l2) => l2.CompareTo(l1));
                        var items = new Dictionary<byte, Item>();
                        var rows = columns.ToDictionary(k => k.Key, v => v.Value.OrderBy(c => c.Item1).ToDictionary(x => x.Item1, y => y.Item2));
                        string col1String = string.Join(Environment.NewLine, columns["1_Code"]);
                        string col2String = string.Join(Environment.NewLine, columns["2_Description"]);
                        string col3String = string.Join(Environment.NewLine, columns["3_Qty"]);
                        Item rollingItem = new Item();
                        foreach (byte lineNbr in lineNbrs)
                        {
                            if (rollingItem.Code == null & rows["1_Code"].ContainsKey(lineNbr))
                                rollingItem.Code = rows["1_Code"][lineNbr];

                            if (rollingItem.Description == null & rows["2_Description"].ContainsKey(lineNbr))
                            {
                                string pad = new string(' ', 8);
                                string description = rows["2_Description"][lineNbr] + pad;
                                string newDescription = description;
                                newDescription = Regex.Replace(newDescription, "\\sTRK\\s", "     ") + pad; // dont know what these are
                                newDescription = Regex.Replace(newDescription, " N[A-Z][0-9]{2} ", pad) + pad; // NS15 or NF13 etc
                                while (Regex.IsMatch(newDescription, " ([A-Z])\\1 "))
                                {
                                    newDescription = Regex.Replace(newDescription, " ([A-Z])\\1 ", pad) + pad; // AA, BB, etc
                                }
                                newDescription = Regex.Replace(newDescription, " {2,}", " ") + pad;
                                newDescription = newDescription.Trim();
                                rollingItem.Description = newDescription;
                            }   

                            if (rollingItem.Qty == 0 & rows["3_Qty"].ContainsKey(lineNbr))
                                rollingItem.Qty = int.Parse(rows["3_Qty"][lineNbr]);

                            if (rollingItem.Complete)
                            {
                                items.Add(lineNbr, rollingItem);
                                rollingItem = new Item();
                            }
                        }
                        items_byPage[pageNbr] = items;
                    }
                    const byte halfway= 50; // 50 is halfway mark
                    const byte indent = 8; // 50 is halfway mark
                    string rightPad = new string(' ', halfway);
                    string leftPad = new string(' ', indent);
                    var all = new List<string>();
                    var line1a = $"Client: {client.Trim()}";
                    var line1b = $"Order: {order.Trim()}";
                    var line1 = (line1a + rightPad).Substring(0, halfway) + (line1b + rightPad).Substring(0, halfway);
                    all.Add((leftPad + line1).Substring(0, halfway * 2));

                    var line2a = $"Contact: {contact.Trim()}";
                    var line2b = $"Phone: {phone.Trim()}";
                    var line2 = (line2a + rightPad).Substring(0, halfway) + (line2b + rightPad).Substring(0, halfway);
                    all.Add((leftPad + line2).Substring(0, halfway * 2));

                    var line3a = $"Rep: {repName.Trim()}";
                    var line3b = $"email: {repEmail.Trim()}";
                    var line3 = (line3a + rightPad).Substring(0, halfway) + (line3b + rightPad).Substring(0, halfway);
                    all.Add((leftPad + line3).Substring(0, halfway * 2));

                    const byte col1 = 12;
                    const byte col2 = 81;
                    const byte col3 = 3;
                    all.Add("┏" + new string('━', col1) + "┳" + new string('━', col2) + "┳" + new string('━', col3) + "┓");
                    all.Add("┃" + ("Item#" + new string(' ', col1)).Substring(0, col1) + "┃" + ("Desc." + new string(' ', col2)).Substring(0, col2) + "┃Qty┃");
                    all.Add("┣" + new string('━', col1) + "╋" + new string('━', col2) + "╋" + new string('━', col3) + "┫");
                    foreach (var item in items_byPage[1])
                    {
                        all.Add("┃" + (item.Value.Code + new string(' ', col1)).Substring(0, col1) + "┃" + (item.Value.Description + new string(' ', col2)).Substring(0, col2) + $"┃" + (item.Value.Qty + new string(' ', col3)).Substring(0, col3) + "┃");
                    }
                    all.Add("┗" + new string('━', col1) + "┻" + new string('━', col2) + "┻" + new string('━', col3) + "┛");
                    string printout = string.Join(Environment.NewLine, all);
                    Debugger.Break();
                }
            }
        }
        private struct Item
        {
            public string Code { get; set; }
            public string Description { get; set; }
            public int Qty { get; set; }
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