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
using HtmlAgilityPack;
using System.Globalization;

namespace fhInventoryEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static readonly DirectoryInfo jobsFolder = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs");
        private static readonly List<DirectoryInfo> validFolders = new List<DirectoryInfo>(jobsFolder.EnumerateDirectories().Where(d => !(d.Name.StartsWith("a_") | d.Name.StartsWith("z_"))));
        private static readonly DirectoryInfo samplesFolder = jobsFolder.CreateSubdirectory("z_samples");
        private static readonly HtmlDocument htmlEditor = new HtmlDocument();
        private static readonly CultureInfo culture = new CultureInfo("en-CA");

        //private static readonly string accountName = "Centre Beaubien"; // CentreLeCap, Centre Beaubien
        //private static readonly DirectoryInfo jobInfo = new DirectoryInfo($"{jobsFolder.FullName}{accountName}\\");
        //private static readonly string htmlPath = $"{jobInfo.FullName}a_jobSummary\\index.html";

        public enum PageRegion { none, contact, table_disclaimer, table_hdr, table_data, footer }
        public enum DocumentType { none, quote, delivery, invoice }
        public enum DocumentLanguage { none, english, french }
        public struct ClientForm
        {

        }
        public struct XXX
        {
            public DocumentType type;
            public DocumentLanguage language;
            public override string ToString() => $"{type} [{language}]";
        }
        internal enum InvoiceSection { none, expense, labour, totals }
        public struct Invoice
        {
            public bool Valid { get; }
            public string Name { get; }
            public FileInfo Path { get; }
            public DateTime Date { get; }
            public DateTime Date_begin { get; }
            public DateTime Date_end { get; }
            public Table Table { get; }
            public double Total_costs { get; }
            public double Total_labour { get; }
            public double Total_hours { get; }
            public double Total { get; }
            public double Hourly_rate { get; }
            public double Total_calc { get; }
            public double Out => Math.Round(Total - Total_calc, 2);
            public bool Balances => Out == 0;
            public string Text { get; }
            public Invoice(FileInfo invoice)
            {
                Valid = false;
                Name = invoice.Name;
                Path = invoice;
                Date = new DateTime();
                Date_begin = new DateTime();
                Date_end = new DateTime();
                Total_costs = 0;
                Total_labour = 0;
                Total_hours = 0;
                Total = 0;
                Hourly_rate = 0;
                Total_calc = 0;
                Text = null;

                Table = new Table();
                var col0 = new Table.Column("path", typeof(string));
                var col1 = new Table.Column("date", typeof(DateTime));
                var col2 = new Table.Column("At", typeof(string));
                var col3 = new Table.Column("type", typeof(string));
                var col4 = new Table.Column("code", typeof(string));
                var col5 = new Table.Column("text", typeof(string));
                var col6 = new Table.Column("qty", typeof(double));
                var col7 = new Table.Column("unit", typeof(double));
                var col8 = new Table.Column("ext_pdf", typeof(double));
                var col9 = new Table.Column("ext_calc", typeof(double));
                Table.Columns.AddRange(new Table.Column[] { col0, col1, col2, col3, col4, col5, col6, col7, col8, col9 });

                try
                {
                    using (PdfDocument document = PdfDocument.Open(invoice.FullName))
                    {
                        var dateFormats = new[] { "MM/d/yyyy", "MM/dd/yyyy", "d-MMM-yy", "dd-MMM-yy", "yyyy-MM-dd" };
                        var section = new InvoiceSection();
                        PdfRectangle typeColumn = default;
                        PdfRectangle gstColumn = default;
                        foreach (Page page in document.GetPages())
                        {
                            var words = new List<Word>(page.GetWords());
                            SortWords(words);
                            var words_byLine = Get_pageWords(page);
                            var text_byLine = words_byLine.ToDictionary(k => k.Key, v => string.Join(" ", v.Value.Select(w => w.Text)));
                            Text = string.Join(Environment.NewLine, text_byLine.Values);
                            if (page.Number == 1)
                            {
                                // first page MUST have "Remit to: Sean Glover", otherwise it is not a pdf invoice
                                if (Text.Contains("Remit to: Sean Glover"))
                                {
                                    DateTime.TryParse(words_byLine[1][1].Text, out DateTime invoiceDate); // second line, second word
                                    Date = invoiceDate;
                                    Valid = true;
                                }
                                else
                                {
                                    //Debugger.Break();
                                    break;
                                }
                            }
                            foreach (var stringOfWords in words_byLine.Values)
                            {
                                string line = string.Join("|", stringOfWords.Select(w => w.Text));
                                Word firstWord = stringOfWords.First();
                                if (stringOfWords.Count == 1)
                                {
                                    if (firstWord.Text.ToLowerInvariant() == "expenses") section = InvoiceSection.expense;
                                    if (firstWord.Text.ToLowerInvariant() == "labour") section = InvoiceSection.labour;
                                }
                                if (line.ToLowerInvariant() == "time|&|materials") section = InvoiceSection.totals;
                                if (DateTime.TryParseExact(firstWord.Text, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                                {
                                    var col2_site = new List<Word>(stringOfWords.Skip(1).Where(w => w.BoundingBox.Right < typeColumn.Left));
                                    var col3_code = new List<Word>(stringOfWords.Skip(1).Where(w => w.BoundingBox.IntersectsWith(typeColumn)));
                                    var col4_desc = new List<Word>(stringOfWords.Skip(1).Where(w => w.BoundingBox.Left > typeColumn.Right & w.BoundingBox.Right < gstColumn.Left));
                                    var location = string.Join(" ", col2_site.Select(w => w.Text));
                                    var code = col3_code.First().Text;
                                    var description = string.Join(" ", col4_desc.Select(w => w.Text));
                                    var numbers = new List<double>();
                                    foreach (var word in stringOfWords.Skip(stringOfWords.Count - 3))
                                    {
                                        var isNumber = double.TryParse(word.Text, NumberStyles.Currency, culture, out double number);
                                        numbers.Add(number);
                                    }
                                    var qty = Math.Round(numbers[0], 2);
                                    var unit = Math.Round(numbers[1], 2);
                                    var ext_pdf = Math.Round(numbers[2], 2);
                                    var ext_calc = Math.Round(qty * unit, 2);
                                    //if (Math.Abs(ext_calc - ext_pdf) >= 2.0) Debugger.Break(); // out more than 2$
                                    Table.Rows.Add(new object[] { invoice.Name, date, location, section.ToString(), code, description, qty, unit, ext_pdf, ext_calc });
                                }
                                else if (typeColumn.Area == 0)
                                {
                                    var typeWords = new List<Word>(stringOfWords.Where(w => w.Text.ToLower() == "type"));
                                    var gstWords = new List<Word>(stringOfWords.Where(w => w.Text.ToLower() == "gst"));
                                    if (typeWords.Any())
                                    {
                                        var typeBounds = typeWords.First().BoundingBox;
                                        var gstBounds = gstWords.First().BoundingBox;
                                        typeColumn = new PdfRectangle(typeBounds.Left, 0, typeBounds.Right, page.CropBox.Bounds.Height);
                                        gstColumn = new PdfRectangle(gstBounds.Left, 0, gstBounds.Right, page.CropBox.Bounds.Height);
                                    }
                                }
                                else if (section == InvoiceSection.totals & line.ToLowerInvariant() != "time|&|materials")
                                {
                                    // summary section
                                    var isNumber = double.TryParse(line.Split('|').Last(), out double total);
                                    if (Regex.IsMatch(line, "total[|]invoice", RegexOptions.IgnoreCase))
                                    {
                                        section = InvoiceSection.none;
                                        Total = total;
                                        break;
                                    }
                                    if (Regex.IsMatch(line, "(total ){0,1}(materials|expenses|labour)", RegexOptions.IgnoreCase))
                                    {
                                        //if (total == 0) Debugger.Break();
                                        if (Regex.IsMatch(line, "expenses|total materials", RegexOptions.IgnoreCase))
                                            Total_costs = total;
                                        else
                                            Total_labour = total;
                                    }
                                }
                            }
                            var rowsExpense = new List<Table.Row>(Table.AsEnumerable.Where(r => r["type"].ToString() == "expense"));
                            var rowsLabour = new List<Table.Row>(Table.AsEnumerable.Where(r => r["type"].ToString() == "labour"));
                            var workDates = new List<DateTime>(rowsLabour.Select(r => (DateTime)r["date"]));
                            if (workDates.Any())
                            {
                                Date_begin = workDates.Min();
                                Date_end = workDates.Max();
                            }
                            if (rowsExpense.Any())
                            {

                            }
                            if (rowsLabour.Any())
                            {
                                Hourly_rate = rowsLabour.Average(r => (double)r["unit"]);
                                Total_hours = rowsLabour.Sum(r => (double)r["qty"]);
                            }
                            Total_calc = Math.Round(Table.AsEnumerable.Sum(r => (double)r["ext_calc"]), 2);
                        }
                    }
                }
                catch (PdfDocumentFormatException ex) { Console.WriteLine(ex.Message); }
                catch (InvalidOperationException ex) { Console.WriteLine(ex.Message); }
            }
            public override string ToString()
            {
                var balances = Balances ? "Y" : "N";
                var overUnder = Balances ? string.Empty : (Out > 0 ? " Over" : " Under") + $" {Out:N2}";
                return $"{Name} Date {Date:yyyy-MM-dd} Total {Total:N2} Balances {balances}{overUnder}";
            }
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

            //Test_all();

            //var mackay = Parse_form(new FileInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\MacKay MTL\\RPOPP7167 Philip E Layton Quote.pdf"));
            //Debugger.Break();
            //var yaldei = Parse_form(new FileInfo("C:/Users/SeanGlover/Desktop/Personal/FH/Jobs/Yaldei/delivery_Yaldei.pdf"));
            //Debugger.Break();

            //Get_pdfForms();

            Edit_html("centrelecap");
            Debugger.Break();

            //var testFolder = Get_invoices(new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\JulesLeger"));
            //var inv = new Invoice(new FileInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Carlyle MTL\\billing\\invoice_Carlyle MTL [0]_2015-06-27.pdf"));


        }
        internal static void Get_stats()
        {
            const byte nbrYears = 5;
            var startTime = DateTime.Now;
            var allTables = new Table();
            var invoices = new List<Invoice>(Get_invoices().Where(i => i.Date.AddYears(nbrYears) >= DateTime.Now));
            invoices.Sort((i1, i2) => i1.Name.CompareTo(i2.Name));
            foreach (var invoice in invoices)
                allTables.Merge(invoice.Table);
            var endTime = DateTime.Now;
            var elapsed = endTime - startTime;
            var summary = string.Join(Environment.NewLine, invoices.Select(i => i.ToString()));

            var expenseRows = new List<Table.Row>(allTables.AsEnumerable.Where(r => r["type"].ToString() == "expense"));
            var labourRows = new List<Table.Row>(allTables.AsEnumerable.Where(r => r["type"].ToString() == "labour"));
            var expenseGroups = new Dictionary<string, List<Table.Row>>();
            foreach (var row in expenseRows)
            {
                var code = row["code"].ToString();
                if (!expenseGroups.ContainsKey(code)) expenseGroups[code] = new List<Table.Row>();
                expenseGroups[code].Add(row);
            }
            var mileage = Math.Round(expenseGroups["T"].Where(r => r["at"].ToString().ToLower() == "mileage").Select(r => (double)r["ext_calc"]).Sum(), 2);
            var expense = Math.Round(expenseRows.Select(r => (double)r["ext_calc"]).Sum(), 2);
            var expGrvy = Math.Round((expense - mileage) * .2, 2) + mileage;
            var hrsTime = Math.Round(labourRows.Select(r => (double)r["qty"]).Sum(), 2);
            var hrsPaid = Math.Round(labourRows.Select(r => (double)r["ext_calc"]).Sum(), 2);
            var ttlRcvd = Math.Round(hrsPaid + expGrvy, 2);
            var ttlYear = Math.Round(ttlRcvd / nbrYears, 2);
            var ttls_ByYear = new Dictionary<int, Dictionary<string, double>>();
            foreach (var row in allTables.AsEnumerable)
            {
                var date = (DateTime)row["date"];
                if (!ttls_ByYear.ContainsKey(date.Year))
                    ttls_ByYear[date.Year] = new Dictionary<string, double> { { "hrs", 0 }, { "car", 0 }, { "exp", 0 } };
                string type = row["type"].ToString().ToLower();
                string subType = type == "expense" ? row["at"].ToString().ToLower() == "mileage" ? "car" : "exp" : "hrs";
                ttls_ByYear[date.Year][subType] += (double)row["ext_calc"] * (subType == "exp" ? .2 : 1);
            }
            ttls_ByYear = ttls_ByYear.OrderByDescending(k => k.Key).Take(5).ToDictionary(k => k.Key, v => v.Value.ToDictionary(kk => kk.Key, vv => Math.Round(vv.Value, 2)));
            var revn_byYear = ttls_ByYear.ToDictionary(k => k.Key, v => Math.Round(v.Value.Values.Sum(), 2));
            var avg_5yrRevn = Math.Round((6000 + revn_byYear.Values.Sum()) / revn_byYear.Count, 2);

            // average
            // 5 year history knowing some records are missing (Opitciwan * 2, etc)
            //  1,124.17 -->    hours worked -225 hours each year or 6 weeks full time
            //  5,434.92 -->    all mileage                             -  1,087 / yr
            //  9,903.27 -->    all expenses gravy (20% of purchased)   -  1,981 / yr
            // 52,930.20 -->    labour paid                             - 10,586 / yr
            // ==========
            // 68,268.39 -->    all earnings                            - 13,654 / yr --> 25k gross
            Debugger.Break();
        }
        internal static void Get_imbalances(bool openFile = true)
        {
            var startTime = DateTime.Now;
            var allTables = new Table();
            var invoices = new List<Invoice>(Get_invoices().Where(i => i.Date.AddYears(7) >= DateTime.Now));
            invoices.Sort((i1, i2) => i1.Name.CompareTo(i2.Name));
            var imbalances = new List<Invoice>(invoices.Where(i => Math.Abs(i.Out) > .5));
            foreach (var invoice in imbalances)
            {
                allTables.Merge(invoice.Table);
                if (openFile) Process.Start(invoice.Path.FullName);
            }
            var endTime = DateTime.Now;
            var elapsed = endTime - startTime;
            var summary = string.Join(Environment.NewLine, imbalances.Select(i => i.ToString()));
            Debugger.Break();
        }
        private static List<Invoice> Get_invoices()
        {
            var invoices = new List<Invoice>();
            foreach (var jobFolder in validFolders)
                invoices.AddRange(Get_invoices(jobFolder));
            return invoices;
        }
        private static List<Invoice> Get_invoices(DirectoryInfo invoiceFolder)
        {
            var invoices = new List<Invoice>();
            var invoiceFiles = new List<FileInfo>(invoiceFolder.EnumerateFiles("*.pdf", SearchOption.AllDirectories));
            if (invoiceFiles.Any())
            {
                foreach (var invoiceFile in invoiceFiles)
                {
                    var invoice = new Invoice(invoiceFile);
                    if (invoice.Valid)
                        invoices.Add(invoice);
                    else { }
                        //Debugger.Break();
                }
            }
            var files = new List<string>();
            int fileIndex = 0;
            DirectoryInfo jobFolder = null;
            foreach (var invoice in invoices.OrderBy(i => i.Date))
            {
                DirectoryInfo folder = invoice.Path.Directory;
                // C:\Users\SeanGlover\Desktop\Personal\FH\Jobs\Cadens Lighthouse <-- validFolders[0]
                // C:\Users\SeanGlover\Desktop\Personal\FH\Jobs\Cadens Lighthouse <-- folder.. same but DirectoryInfo IEquatable doesn't work they way you'd expect
                var folders = new List<string>(validFolders.Select(di => di.FullName));
                while (!folders.Contains(folder.FullName)) folder = folder.Parent;
                var billFolder = folder.CreateSubdirectory("billing"); // no harm even if it already exists
                jobFolder = folder;
                string guidPath = $"{billFolder}\\{Guid.NewGuid()}.pdf";
                string newIndex = invoices.Count == 1 ? string.Empty : $" [{fileIndex}]";
                string newPath = $"{billFolder}\\invoice_{folder.Name}{newIndex}_{invoice.Date:yyyy-MM-dd}.pdf";
                File.Move(invoice.Path.FullName, guidPath);
                File.Move(guidPath, newPath);
                //files.Add($"date {v.Item2:yyyy-MM-dd} start {v.Item3:yyyy-MM-dd} end {v.Item4:yyyy-MM-dd} MAT {v.Item5:N2} LBR {v.Item6:N2} TTL {v.Item7:N2}");
                fileIndex++;
            }
            if (files.Any())
                File.WriteAllText($"{jobsFolder}\\invoice_{jobFolder.Name}.txt", string.Join(Environment.NewLine, files));
            //Debugger.Break();
            return invoices;
        }
        private static Dictionary<byte, List<Word>> Get_pageWords(Page page)
        {
            if (page == null) return null;
            var lines = new Dictionary<byte, List<Word>>();
            var indexedWords = new Dictionary<int, Word>();
            var words_byLine = new Dictionary<int, List<Word>>();
            var words = new List<Word>(page.GetWords());
            SortWords(words);
            foreach (var word in words) indexedWords[indexedWords.Count] = word;
            byte lineCounter = 0;
            var startTime = DateTime.Now;
            var timedOut = false;
            while (indexedWords.Count > 0 & !timedOut)
            {
                timedOut = DateTime.Now.Subtract(startTime).TotalSeconds >= 3;
                var wordsByLine = new List<KeyValuePair<int, Word>>
                    (indexedWords.Where(w => !(from x in indexedWords where w.Value.BoundingBox.Top < x.Value.BoundingBox.Bottom select x).Any()));
                if (!wordsByLine.Any())
                    break;
                words_byLine[lineCounter] = new List<Word>();
                foreach (var word in wordsByLine)
                {
                    words_byLine[lineCounter].Add(word.Value);
                    indexedWords.Remove(word.Key);
                }
                words_byLine[lineCounter].Sort((w1, w2) => w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left));
                lines[lineCounter] = words_byLine[lineCounter];
                lineCounter++;
            }
            if (timedOut) Debugger.Break();
            return lines;
        }
        internal static Dictionary<byte, string> Get_pageText(Page page)
        {
            var get_pageWords = Get_pageWords(page);
            return get_pageWords == null ? null : get_pageWords.ToDictionary(k => k.Key, v => string.Join(" ", v.Value.Select(w => w.Text)));
        }
        internal void Edit_html(string folder, string contactTitle = "Program and services coordinator")
        {
            // get last delivery form in a folder and tie it back to the matching quote AND invoice
            // 
            var directory = Get_directoryByName(folder);
            if (directory != null)
            {
                var save = SaveAll_byType(directory);
                var forms = Parse_forms(directory);
                var invoices = Get_invoices(directory);
                if (forms.Item3.Any())
                {
                    string sourceIndexHTML = "C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\a_jobSummary\\index.html";
                    string destinationIndexHTML = $"{directory.FullName}\\a_jobSummary\\index.html";
                    bool createFiles = !File.Exists(destinationIndexHTML);
                    if (createFiles)
                        DirectoryCopy("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\a_jobSummary\\", $"{directory.FullName}\\a_jobSummary\\", true);
                    else
                        File.Copy(sourceIndexHTML, destinationIndexHTML, true);

                    string sourceHtml = File.ReadAllText(sourceIndexHTML);
                    htmlEditor.LoadHtml(sourceHtml);

                    var deliveryFiles = new List<FileInfo>(directory.EnumerateFiles("*delivery_*.pdf", SearchOption.TopDirectoryOnly));
                    if (deliveryFiles.Any())
                    {
                        deliveryFiles.Sort((d1, d2) => d2.CreationTime.CompareTo(d1.CreationTime)); // last delivery
                        var deliveryForm = Parse_form(deliveryFiles.Last());
                        var orders = new List<Table.Row>(deliveryForm.Item3.AsEnumerable.Where(r => r["key"].ToString() == "Order#"));
                        string orderNbr = orders.Any() ? orders[0]["value"].ToString() : string.Empty;

                        // P0810161 - NSOPP5357, Order#      |P0810161 - revised #2
                        var quoteFiles = new List<FileInfo>(directory.EnumerateFiles("*quote_*.pdf", SearchOption.TopDirectoryOnly));
                        quoteFiles.Sort((d1, d2) => d2.CreationTime.CompareTo(d1.CreationTime)); // last quote
                        var quoteForms = new List<Tuple<Table, XXX, Table>>();
                        foreach (var quoteFile in quoteFiles)
                            quoteForms.Add(Parse_form(quoteFile));

                        #region" html data "
                        /// there are 6 elements sourced from a delivery form, but only 4 are used in the web page ( rep and rep's email are not used )
                        /// order#, customer, contact, and phone
                        var businessName = ProperCase(deliveryForm.Item3.AsEnumerable.Where(r => r["key"].ToString() == "Customer").Select(r => r["value"].ToString()).First());
                        var contact_name = ProperCase(deliveryForm.Item3.AsEnumerable.Where(r => r["key"].ToString() == "Contact").Select(r => r["value"].ToString()).First());
                        var contact_phone = deliveryForm.Item3.AsEnumerable.Where(r => r["key"].ToString() == "Phone#").Select(r => r["value"].ToString()).First();
                        // ----------- these elements come from a quote form
                        var address_street = string.Empty;
                        var address_city = string.Empty;
                        var address_province = string.Empty;
                        var address_postal = string.Empty;
                        var contact_email = string.Empty;
                        var quoteOrders = new List<Tuple<Table, XXX, Table>>(quoteForms.Where(qf => qf.Item3.AsEnumerable.Where(r => r["key"].ToString() == "Order#" & r["value"].ToString().Contains(orderNbr)).Any()));
                        if (quoteOrders.Any())
                        {
                            var contactTable = new Table();
                            foreach (var quoteOrder in quoteOrders) contactTable.Merge(quoteOrder.Item3);
                            address_street = ProperCase(string.Join(" ", contactTable.AsEnumerable.Where(r => r["key"].ToString().StartsWith("Address")).Select(r => r["value"].ToString())));
                            address_city = ProperCase(contactTable.AsEnumerable.Where(r => r["key"].ToString() == "City").Select(r => r["value"].ToString()).First());
                            address_province = provinces.Where(p => p.Item1 == ProperCase(contactTable.AsEnumerable.Where(r => r["key"].ToString() == "Province").Select(r => r["value"].ToString()).First())).First().Item3;
                            address_postal = ProperCase(contactTable.AsEnumerable.Where(r => r["key"].ToString() == "PostalCode").Select(r => r["value"].ToString()).First());
                            contact_email = ProperCase(contactTable.AsEnumerable.Where(r => r["key"].ToString() == "email-client").Select(r => r["value"].ToString()).First());
                        }
                        else
                        {
                            /// no matching order for the delivery form
                            /// confirmed for CentreLeCap-
                            /// no connection between the Quote pdf  "Reference #: P0914839"
                            /// and the Delivery pdf                 "Order:	P0917654 / P0927227"
                            /// but the two ARE connected as the equipment on the delivery closely matches the quote

                            var hitRates = new List<Tuple<double, ClientForm>>();
                            foreach (var quoteForm in quoteForms)
                            {
                                var codesQuoted = new List<string>(quoteForm.Item1.Columns["item"].Values.Select(i => i.Value.ToString()));
                                var codesDelivered = new List<string>(deliveryForm.Item1.Columns["item"].Values.Select(i => i.Value.ToString()));
                                var codesCommon = new List<string>(codesQuoted.Intersect(codesDelivered));
                                var hitRate = new double[] { 1, codesCommon.Count }.Max() / new double[] { codesQuoted.Count, codesDelivered.Count }.Min();
                                Debugger.Break();
                            }
                            Debugger.Break();

                        }
                        Debugger.Break();
                        #endregion
                        htmlEditor.GetElementbyId("documentDate").SetAttributeValue("value", $"{DateTime.Now:yyyy-MM-dd}");
                        htmlEditor.GetElementbyId("businessName").SetAttributeValue("value", businessName);
                        htmlEditor.GetElementbyId("purchaseOrder").SetAttributeValue("value", orderNbr);
                        htmlEditor.GetElementbyId("contactName").SetAttributeValue("value", contact_name);
                        htmlEditor.GetElementbyId("contactPhone").SetAttributeValue("value", contact_phone);
                        htmlEditor.GetElementbyId("contactEmail").SetAttributeValue("value", contact_email);
                        htmlEditor.GetElementbyId("contactTitle").SetAttributeValue("value", contactTitle);
                        htmlEditor.GetElementbyId("businessName").SetAttributeValue("value", businessName);
                        htmlEditor.GetElementbyId("addressStreet").SetAttributeValue("value", address_street);
                        htmlEditor.GetElementbyId("addressCity").SetAttributeValue("value", address_city);
                        var provs = new List<HtmlNode>(htmlEditor.GetElementbyId("Provinces").ChildNodes);
                        var mkys = new List<HtmlNode>();
                        foreach (var option in provs)
                            option.Attributes.Remove("selected");
                        htmlEditor.GetElementbyId(address_province).SetAttributeValue("selected", "\"\"");
                        htmlEditor.GetElementbyId("addressPostalCode").SetAttributeValue("value", address_postal);
                    }
                    using (StreamWriter sw = new StreamWriter(destinationIndexHTML))
                        htmlEditor.Save(sw);
                    Process.Start(destinationIndexHTML);

                    // 2] update the products info
                    //Debugger.Break();
                }
            }
        }
        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
                throw new DirectoryNotFoundException("Source directory does not exist or could not be found: " + sourceDirName);

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName))
                Directory.CreateDirectory(destDirName);

            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
        }
        internal static void Get_samples()
        {
            var samples = Get_DocumentTypes();
            samples.Remove(DocumentType.none);
            foreach (var sampleType in samples)
                foreach (var pdf in sampleType.Value)
                    File.Copy(pdf.FullName, $"{samplesFolder}{pdf.Name}", true);
        }
        internal static List<FileInfo> Get_pdfForms(DocumentType docType = DocumentType.delivery, bool openFile = true)
        {
            var pdfForms = new List<FileInfo>();
            foreach (var jobFolder in validFolders)
                foreach (var jobFile in jobFolder.EnumerateFiles($"*{docType}_*.pdf"))
                {
                    pdfForms.Add(jobFile);
                    if (openFile) Process.Start(jobFile.FullName);
                }
            return pdfForms;
        }
        internal static void Test_all()
        {
            //var form = Parse_form(new FileInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Centre Evasion/quote_Centre Evasion [0].pdf"));
            //Debugger.Break();

            DateTime startTime = DateTime.Now;
            var forms = Parse_forms();

            string html = forms.Item1.HTML;
            string contactData = string.Join(Environment.NewLine, forms.Item2.Lines);

            var codeList = new List<string>(forms.Item1.Rows.Values.Select(r => r["item"].ToString()).Distinct());
            codeList.Sort();
            string codes = string.Join(Environment.NewLine, codeList);
            DateTime endTime = DateTime.Now;
            TimeSpan elapsed = endTime - startTime;
            Debugger.Break();
        }
        internal void Testing()
        {
            //var form = Parse_form(new FileInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Prescott Russell_Hawkesbury/delivery_Prescott Russell_Hawkesbury.pdf"));
            //Debugger.Break();
            //new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\Centre Evasion\\")

            DateTime startTime = DateTime.Now;
            var forms = Parse_forms();
            string contacts = string.Join(Environment.NewLine, forms.Item2.AsEnumerable.Select(r => string.Join("|", r["value"].ToString())));
            string html = forms.Item1.HTML;
            var codeList = new List<string>(forms.Item1.Rows.Values.Select(r => r["item"].ToString()).Distinct());
            codeList.Sort();
            string codes = string.Join(Environment.NewLine, codeList);
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
            foreach (var jobFolder in validFolders)
            {
                var tableFiles = new List<FileInfo>(jobFolder.EnumerateFiles("*.pdf", SearchOption.TopDirectoryOnly));
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
        }
        internal void Addresses()
        {
            var stringOfAddresses = "Dolbeau, Mistassini, Québec G8L 2V4\r\nDolbeau Mistassini, QC G8L 2V4\r\nOTTAWA, ON K1Z 6R8 CANADA\r\nMONTREAL, QC H4B 1B7 CANADA\r\nOTTAWA, ON K1L 8H2 CANADA\r\nL'ORIGNAL, ON K0B 1K0 CANADA\r\nDOLBEAU MISTASSINI, QC G8L 2V4\r\nOTTAWA, ON K1Z 6R8 C\r\nMONTREAL, QC H3T 1B1 CANADA\r\nMONTREAL, QC H3R 2G3\r\nMONTREAL, QC H3W 1J6 CANADA";
            List<string> addresses = new List<string>(stringOfAddresses.Split(new string[] { Environment.NewLine }, StringSplitOptions.None));
            //addresses.Sort((a1, a2) => a2.CompareTo(a1));
            var properAddresses = new List<string>();
            foreach (var address in addresses)
            {
                string testValue = address; //addresses[0];
                var addressDict = CityProvincPostal(testValue);
                properAddresses.Add(string.Join("|", addressDict.Values));
            }
            Debugger.Break();

            DirectoryInfo di = new DirectoryInfo("C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\xxx\\");
            foreach (FileInfo file in di.EnumerateFiles("*.txt"))
                addresses.Add(File.ReadAllText(file.FullName));
            string allAddresses = string.Join("\n", addresses.Distinct());
            Debugger.Break();
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
                matchedDirectories.AddRange(jobsFolder.EnumerateDirectories().Where(d => Regex.IsMatch(d.Name, foldername_Or_Filename, RegexOptions.IgnoreCase)));

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
            foreach (var jobFolder in validFolders)
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
            foreach (var folder in validFolders)
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
        private static XXX Get_documentTypeLanguage(FileInfo pdfinfo)
        {
            string[] keyDeliveryWords = new string[]
            {
                "Formulaire de vérification de livraison",
                "Les commandes doivent être inspectées pour dommages d'expédition",
                "Shipping damage claims will only be accepted by",
                "Delivery Verification Form"
            };
            var documentPages = Get_pdfText(pdfinfo);
            if (documentPages == null) return new XXX() { language = DocumentLanguage.none, type = DocumentType.none };
            var documentText = string.Join(Environment.NewLine, documentPages.Values);
            
            // quote match MUST be first as the quote contains the delivery disclaimer statement, but delivery forms DON'T have QUOTE#/SOUMISSION
            var quoteMatch = Regex.Match(documentText, "(QUOTE|SOUMISSION) {0,}#:", RegexOptions.IgnoreCase);
            if (quoteMatch.Success)
            {
                var quoteLanguage = Regex.IsMatch(quoteMatch.Value, "SOUMISSION {0,1}#:") ? DocumentLanguage.french : DocumentLanguage.english;
                return new XXX() { language = quoteLanguage, type = DocumentType.quote };
            }
            
            var dlvryMatch = Regex.Match(documentText, $"({string.Join("|", keyDeliveryWords)})", RegexOptions.IgnoreCase);
            if (dlvryMatch.Success)
            {
                var deliveryLanguage = Regex.IsMatch(dlvryMatch.Value, $"({string.Join("|", keyDeliveryWords.Take(2))})") ? DocumentLanguage.french : DocumentLanguage.english;
                return new XXX() { language = deliveryLanguage, type = DocumentType.delivery };
            }

            var noneLanguage = Regex.IsMatch(documentText, "[àâçéèêëîïôûùüÿñæœ]", RegexOptions.IgnoreCase) ? DocumentLanguage.french : DocumentLanguage.english;
            return new XXX() { language = noneLanguage, type = DocumentType.none };
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
        internal static Tuple<Table, Table, List<FileInfo>> Parse_forms(bool openFiles = false)
        {
            Table allTables = new Table();
            Table allContacts = new Table();
            List<FileInfo> files = new List<FileInfo>();
            foreach (var jobFolder in validFolders)
            {
                var parsedForms = Parse_forms(jobFolder, openFiles);
                allTables.Merge(parsedForms.Item1);
                allContacts.Merge(parsedForms.Item2);
            }
            return Tuple.Create(allTables, allContacts, files);
        }
        private static Tuple<Table, Table, List<FileInfo>> Parse_forms(DirectoryInfo jobFolder, bool openFiles = false)
        {
            Table folderTable = new Table();
            Dictionary<string, List<string>> contacts = new Dictionary<string, List<string>>();
            Table contactsTable = new Table();
            List<FileInfo> files = new List<FileInfo>();
            foreach (var file in jobFolder.EnumerateFiles("*.pdf").Where(f=>Regex.IsMatch(f.Name, "(delivery|quote)_[^.]{1,}\\.pdf")))
            {
                var parsedForm = Parse_form(file, openFiles);
                folderTable.Merge(parsedForm.Item1);
                files.Add(file);
                contactsTable.Merge(parsedForm.Item3);
            }
            return Tuple.Create(folderTable, contactsTable, files);
        }
        private static Tuple<Table, XXX, Table> Parse_form(FileInfo jobinfo, bool openFile = false)
        {
            if (jobinfo == null) return Tuple.Create(new Table(), new XXX(), new Table());
            else
            {
                if (openFile) Process.Start($"{jobinfo.FullName}");

                var ColumnNames = new List<string>();

                Table itemTable = new Table();
                var item = itemTable.Columns.Add("item", typeof(string));
                var desc = itemTable.Columns.Add("desc", typeof(string));
                var qty = itemTable.Columns.Add("qty", typeof(double));
                var path = itemTable.Columns.Add("path", typeof(string));
                var docDate = itemTable.Columns.Add("date", typeof(DateTime));
                itemTable.PrimaryKeys = new Table.Column[] { item };

                var contacts = new Dictionary<string, string>();
                var tableRows = new Dictionary<byte, string>();
                var documentTypeLanguage = Get_documentTypeLanguage(jobinfo);
                var letters_byPage = new Dictionary<byte, Dictionary<byte, Dictionary<double, Dictionary<int, Letter>>>>();
                var consecutiveLetters_byPage = new Dictionary<byte, Dictionary<double, Dictionary<byte, Dictionary<int, Letter>>>>();

                using (PdfDocument document = PdfDocument.Open(jobinfo.FullName))
                {
                    var colRects = new Dictionary<string, PdfRectangle>();
                    var colNames = new Dictionary<string, Word>();
                    var cols = new List<double>();
                    PageRegion pgRegion = new PageRegion();

                    foreach (Page page in document.GetPages())
                    {
                        byte pageNbr = Convert.ToByte(page.Number);
                        string pageText = page.Text;
                        const byte margin = 2;

                        var consecutiveLetters_thisPage = new Dictionary<double, Dictionary<byte, Dictionary<int, Letter>>>();
                        consecutiveLetters_byPage[pageNbr] = consecutiveLetters_thisPage;
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

                        #region" words_thisPage "
                        List<Word> words = new List<Word>(page.GetWords());
                        SortWords(words);
                        var wordDict = new Dictionary<int, Word>();
                        foreach (var word in words)
                            wordDict.Add(wordDict.Count, word);
                        var yGroups = new Dictionary<double, List<Word>>();
                        while (wordDict.Any())
                        {
                            var anyWord = wordDict.First();
                            var word = anyWord.Value;
                            // it is possible to have the word rectangle not include the letters that make up the word!! FFS
                            double bot = new double[] { word.Letters.Min(l => l.GlyphRectangle.Bottom), word.BoundingBox.Bottom }.Min();
                            double top = new double[] { word.Letters.Max(l => l.GlyphRectangle.Top), word.BoundingBox.Top }.Max();
                            double lft = new double[] { word.Letters.Min(l => l.GlyphRectangle.Left), word.BoundingBox.Left }.Min();
                            double rgt = new double[] { word.Letters.Max(l => l.GlyphRectangle.Right), word.BoundingBox.Right }.Max();
                            var rect = new PdfRectangle(new PdfPoint(lft, bot), new PdfPoint(rgt, top));
                            var rectBottom = rect.Bottom;
                            var line = new PdfRectangle(new PdfPoint(0, rectBottom - margin), new PdfPoint(10000, rect.Top + margin));
                            var wordsInLine = wordDict.Where(w => line.IntersectsWith(w.Value.BoundingBox)).ToDictionary(k => k.Key, v => v.Value);

                            //Y = lettersInLine.Min(l => l.Value.GlyphRectangle.Bottom);
                            //Y = LineNbr_round(Y);

                            yGroups[rectBottom] = new List<Word>(wordsInLine.Values);
                            yGroups[rectBottom].Sort((l1, l2) => l1.BoundingBox.Left.CompareTo(l2.BoundingBox.Left));
                            foreach (var wrd in wordsInLine) wordDict.Remove(wrd.Key);
                        }

                        var words_thisPage = new Dictionary<double, List<Word>>();
                        var wordLines_thisPage = new Dictionary<double, string>();
                        foreach (var yGroup in yGroups)
                        {
                            var orderedWords = yGroup.Value;
                            orderedWords.Sort((w1, w2) => w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left));
                            double lineNbr = LineNbr_round(yGroup.Key);
                            words_thisPage[lineNbr] = new List<Word>(orderedWords);
                            wordLines_thisPage[lineNbr] = string.Join(" ", orderedWords.Select(w => w.Text));
                        }
                        #endregion
                        #region" letters_thisPage "
                        List<Letter> letters = new List<Letter>(page.Letters);
                        letters.Sort((l1, l2) =>
                        {
                            int lvl1 = l2.GlyphRectangle.Bottom.CompareTo(l1.GlyphRectangle.Bottom);
                            if (lvl1 != 0) return lvl1;
                            int lvl2 = l1.GlyphRectangle.Left.CompareTo(l2.GlyphRectangle.Left);
                            return lvl2;
                        });
                        var letterDict = new Dictionary<int, Letter>();
                        foreach (var letter in letters)
                            letterDict.Add(letterDict.Count, letter);
                        var Ygroups = new Dictionary<double, List<Letter>>();

                        while (letterDict.Any())
                        {
                            var anyLetter = letterDict.First();
                            PdfRectangle rect;
                            var wordLetters = new List<Word>(words.Where(w => (from x in w.Letters where x == anyLetter.Value select x).Any()));
                            if (wordLetters.Any())
                            {
                                // it is possible to have the word rectangle not include the letters that make up the word!! FFS
                                var word = wordLetters.First();
                                double bot = new double[] { word.Letters.Min(l => l.GlyphRectangle.Bottom), word.BoundingBox.Bottom }.Min();
                                double top = new double[] { word.Letters.Max(l => l.GlyphRectangle.Top), word.BoundingBox.Top }.Max();
                                double lft = new double[] { word.Letters.Min(l => l.GlyphRectangle.Left), word.BoundingBox.Left }.Min();
                                double rgt = new double[] { word.Letters.Max(l => l.GlyphRectangle.Right), word.BoundingBox.Right }.Max();
                                rect = new PdfRectangle(new PdfPoint(lft, bot), new PdfPoint(rgt, top));
                            }
                            else
                                rect = anyLetter.Value.GlyphRectangle;
                            var rectBottom = rect.Bottom;
                            var line = new PdfRectangle(new PdfPoint(0, rectBottom - margin), new PdfPoint(10000, rect.Top + margin));
                            var lettersInLine = letterDict.Where(l => line.IntersectsWith(l.Value.GlyphRectangle)).ToDictionary(k => k.Key, v => v.Value);

                            //Y = lettersInLine.Min(l => l.Value.GlyphRectangle.Bottom);
                            //Y = LineNbr_round(Y);

                            Ygroups[rectBottom] = new List<Letter>(lettersInLine.Values);
                            Ygroups[rectBottom].Sort((l1, l2) => l1.StartBaseLine.X.CompareTo(l2.StartBaseLine.X));
                            foreach (var letter in lettersInLine) letterDict.Remove(letter.Key);
                        }
                        var letters_thisPage = new Dictionary<double, List<Letter>>();
                        var letterLines_thisPage = new Dictionary<double, string>();
                        foreach (var Ygroup in Ygroups)
                        {
                            var orderedLetters = Ygroup.Value;
                            orderedLetters.Sort((l1, l2) => l1.StartBaseLine.X.CompareTo(l2.StartBaseLine.X));
                            double lineNbr = LineNbr_round(Ygroup.Key);
                            letters_thisPage[lineNbr] = new List<Letter>(orderedLetters);

                            byte wordIndex = 0;
                            var avgLetterWidth = orderedLetters.Average(l => l.GlyphRectangle.Width);
                            var remainingLetters = new Dictionary<int, Letter>();
                            foreach (var letter in orderedLetters) remainingLetters.Add(remainingLetters.Count, letter);
                            var firstInChain = remainingLetters.First();
                            consecutiveLetters_thisPage[lineNbr] = new Dictionary<byte, Dictionary<int, Letter>>();
                            var words_thisLine = consecutiveLetters_thisPage[lineNbr];
                            while (remainingLetters.Any())
                            {
                                if (!words_thisLine.ContainsKey(wordIndex))
                                    words_thisLine[wordIndex] = new Dictionary<int, Letter>();
                                remainingLetters.Remove(firstInChain.Key);
                                words_thisLine[wordIndex].Add(words_thisLine[wordIndex].Count, firstInChain.Value);
                                var nextInChain = new List<KeyValuePair<int, Letter>>(remainingLetters.Where(nextLetter => (nextLetter.Value.StartBaseLine.X - firstInChain.Value.EndBaseLine.X) < 2));
                                if (nextInChain.Any())
                                    firstInChain = nextInChain.First();
                                else if (remainingLetters.Any())
                                {
                                    firstInChain = remainingLetters.First();
                                    wordIndex++;
                                }
                            }
                            letterLines_thisPage[lineNbr] = string.Join("■", words_thisLine.Select(w => string.Join("", w.Value.Values.Select(l => l.Value))));
                        }
                        #endregion

                        #region" table is on this page - get column names, and column left positions + contact info ONCE "
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
                            thinnestVerticalLine = Math.Ceiling(thinnestVerticalLine);
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
                            foreach (var vLine in pdfRects_linesVertical.Where(l => Math.Ceiling(l.Width) == thinnestVerticalLine))
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

                            // [663, Date:■le 28 octobre 2016]
                            // [660,  ]
                            // [648, QUOTE#:■NSOPP4361■Consultant:■Nicole■Segal]
                            // [645,  ■ ]
                            // [633, Reference #:■P0701984 modifié-PHASE 1  (Offre spéciale fin d'annéeE)-Mail:  ■nicole.segal@flaghouse.com] <-- letters works better than words
                            // [612, Attention:■Dominique Lambert■Phone #:■418- 276-5101]
                            // [600, Client:■Ecole Ste Therese■Fax #:]
                            // [591, Address:■242, 3 e Avenue■E-Mail:■LambertD@cspaysbleuets.qc.ca]
                            // [582, Dolbeau Mistassini,  QC  G8L 2V4]

                            var wordsContact = new List<Word>();
                            if (documentTypeLanguage.type == DocumentType.delivery)
                            {
                                double VERIFICATION = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "v(e|é)rification", RegexOptions.IgnoreCase))).FirstOrDefault().BoundingBox.Bottom;
                                double PLEASENOTE = new List<Word>(wordsAboveTable.Where(w => Regex.IsMatch(w.Text, "NOTE(R){0,1}:"))).FirstOrDefault().BoundingBox.Top;
                                var wordsAboveDisclaimer = new List<Word>(wordsAboveTable.Where(w => w.BoundingBox.Bottom > PLEASENOTE));
                                wordsContact.AddRange(wordsAboveDisclaimer.Where(w => w.BoundingBox.Top < VERIFICATION));
                                SortWords(wordsContact);

                                #region" define two vertical lines in the contact region - leftside=right of colon, rightside=left of colon"
                                var colons = new List<Word>(wordsContact.Where(w => w.Text.Contains(":")));
                                var colonsRight = new List<Word>(colons.Where(w => (from x in colons where w.BoundingBox.Left > x.BoundingBox.Right select x).Any()));
                                var colonsLeft = new List<Word>(colons.Except(colonsRight));
                                
                                var custContactRep_right = colonsLeft.Max(w => w.BoundingBox.Right);
                                var orderPhoneEmail_left = colonsRight.Min(w => w.BoundingBox.Left);
                                colons.AddRange(wordsContact.Where(w => w.BoundingBox.Left < custContactRep_right)); // FH Rep: or Flaghouse Rep: will include FH/Flaghouse
                                colons = colons.Distinct().ToList();
                                #endregion

                                #region" line1 - top "
                                var line1 = wordsContact.Where(w => w.Text.Trim().ToLower().Contains("client") | w.Text.Trim().ToLower().Contains("customer")).First().BoundingBox;
                                var wordsLine1 = new List<Word>(wordsContact.Where(w => w.BoundingBox.Bottom >= line1.Bottom));
                                var wordsLine1Right = new List<Word>(wordsLine1.Where(w => w.BoundingBox.Left > orderPhoneEmail_left).Except(colons));
                                var wordsLine1Left = new List<Word>(wordsLine1.Except(wordsLine1Right).Except(colons));
                                //------------------------------
                                string customer = string.Join(" ", wordsLine1Left.Select(w => w.Text)).Trim();
                                string order = string.Join(" ", wordsLine1Right.Select(w => w.Text)).Trim();
                                #endregion
                                
                                foreach (var line1Word in wordsLine1) wordsContact.Remove(line1Word);
                                
                                #region" line3 - bottom "
                                SortWords(wordsContact);
                                var wordsLines_2_3 = new List<Word>(wordsContact.Except(colons));
                                var wordsLine3 = new List<Word>(wordsLines_2_3.Where(w => (from x in wordsLines_2_3 where w.BoundingBox.Top < x.BoundingBox.Bottom select x).Any()));
                                var wordsLine3Right = new List<Word>(wordsLine3.Where(w => w.BoundingBox.Left > orderPhoneEmail_left).Except(colons));
                                var wordsLine3Left = new List<Word>(wordsLine3.Where(w => w.BoundingBox.Left > custContactRep_right).Except(wordsLine3Right).Except(colons));
                                //------------------------------
                                string rep = string.Join(" ", wordsLine3Left.Select(w => w.Text)).Trim();
                                string email = string.Join(" ", wordsLine3Right.Select(w => w.Text)).Trim();
                                #endregion

                                foreach (var line1Word in wordsLine3.Union(colons)) wordsContact.Remove(line1Word);

                                #region" line 2 - middle "
                                var wordsLine2Right = new List<Word>(wordsContact.Where(w => w.BoundingBox.Left > orderPhoneEmail_left));
                                var wordsLine2Left = new List<Word>(wordsContact.Except(wordsLine2Right));
                                //------------------------------
                                string contact = string.Join(" ", wordsLine2Left.Select(w => w.Text)).Trim();
                                string phone = string.Join(" ", wordsLine2Right.Select(w => w.Text)).Trim();
                                #endregion

                                contacts = new Dictionary<string, string>
                                {
                                    {"Customer", customer },
                                    {"Order#", order },
                                    {"Contact", contact },
                                    {"Phone#", phone },
                                    {"Rep", rep },
                                    {"email", email },
                                };
                                //File.WriteAllText($"C:\\Users\\SeanGlover\\Desktop\\Personal\\FH\\Jobs\\{Guid.NewGuid()}_contacts.txt", JsonConvert.SerializeObject(contacts, Formatting.Indented));
                                //Debugger.Break();
                                //contacts = Get_contactKeysValues(wordsContact, col3_qtyLeft, 3); // <-- delivery forms have values floating above the line, so more margin needed
                                //var invalidKeys = new List<string>(contacts.Keys.Where(c => c.StartsWith("Address")));
                                //foreach (var invalidKey in invalidKeys) contacts.Remove(invalidKey);
                                //contacts.Remove("email-client");
                            }
                            if (documentTypeLanguage.type == DocumentType.quote)
                            {
                                // contact words / letters are below the phone #
                                // 235 Yorkland Blvd., Suite 105, North York, Ontario, M2J 4Y8
                                // Phone: 1-800-265-6900 Fax: 1-800-265-6922 www.snoezeleninfo.com
                                var phoneNumber = new List<Word>(words.Where(w => w.Text.Contains("1-800-265-6900")));
                                if (phoneNumber.Any())
                                {
                                    double phoneBottom = LineNbr_round(phoneNumber.First().BoundingBox.Bottom);
                                    wordsContact.AddRange(wordsAboveTable.Where(w => w.BoundingBox.Top < phoneBottom));
                                    contacts = Get_contactKeysValues(wordsContact, col3_qtyLeft, 0);

                                    // below: tried extracting via <Letter> but it too reflects some data with letters out of order- corrupted by overlapping
                                    // var cl = consecutiveLetters_thisPage.Where(w => w.Key > tableDataTop & w.Key < phoneBottom).ToDictionary(k => k.Key, v => v.Value);
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
                        #endregion

                        #region" could be any region on the page- as long as there are words (some pages have pictures only) "
                        foreach (var line in words_thisPage.Where(w => w.Value.Any()))
                        {
                            var wordsThisLine = new List<Word>(line.Value);
                            string lineData = string.Join(" ", wordsThisLine.Select(w => w.Text));
                            if (Regex.IsMatch(lineData, "ITEM {0,}# {0,}DESCRIPTION {0,}Q"))
                            {
                                var colWord = colNames["ITEM"];
                                var colRect = colWord.BoundingBox;
                                var boundingRects = new List<PdfRectangle>(pdfRects.Where(r => r.Contains(colRect)).OrderBy(r => r.Area)); // smallest to largest
                                var tableHeadRect = boundingRects.FirstOrDefault(); // this should be the table heading rectangle that contains the column names
                                var intersectRects = new List<PdfRectangle>(pdfRects.Where(r => tableHeadRect.IntersectsWith(r)).OrderBy(r => r.Width));
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
                                var firstWord = wordsThisLine[0];
                                Letter firstLetter_inWord = firstWord.Letters[0];
                                var colIndex = cols.IndexOf(cols.Where(c => c < firstLetter_inWord.StartBaseLine.X).Max());
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
                                    itemTable.Rows.Add(new object[] { cell1_item, cell2_desc, qtyCell, jobinfo.Name, jobinfo.CreationTime });
                                } // get all the words in a table row once (for column 0)

                                // [117.5] SIGNATURE ■[336.8] DATE
                                bool isTableEnd = lineData.Contains("SIGNATURE") & lineData.Contains("DATE");
                                if (lineData.Contains("NOM IMPRIMÉ") | lineData.Contains("PRINTED NAME")) isTableEnd = true;
                                if (lineData.Contains("VEUILLEZ ENVOYER")) isTableEnd = true;
                                if (lineData.Contains("ONCE COMPLETED")) isTableEnd = true;
                                //if (Regex.IsMatch(lineData, "page [0-9] of [0-9]", RegexOptions.IgnoreCase)) isTableEnd = true;
                                if (lineData.Contains("TERMES ET CONDITIONS")) isTableEnd = true;
                                if (lineData.Contains("TERMS AND CONDITIONS")) isTableEnd = true;
                                if (lineData.Contains("VIRTUAL INSTALLATION AVAILABLE")) isTableEnd = true;
                                if (!lineData.Any()) isTableEnd = true;
                                if (isTableEnd)
                                    pgRegion = PageRegion.footer;
                            }
                        }
                        #endregion
                    }
                    #region" save .txt file - MUST assume pdf filename is the correct format "
                    itemTable.Name = $"{documentTypeLanguage.type}_{jobinfo.Directory.Name}";
                    // save the contacts dictionary + products table in a Tuple as -->  delivery_Centre Evasion
                    string directoryFullPath = $"{jobinfo.DirectoryName}\\{jobinfo.Directory.Name}";
                    var newFilePath = new FileInfo(jobinfo.FullName.Replace(".pdf", ".txt"));
                    var fileTuple = Tuple.Create(contacts, itemTable);
                    File.WriteAllText(newFilePath.FullName, JsonConvert.SerializeObject(fileTuple, Formatting.Indented));
                    #endregion
                }
                var contactsTable = new Table() { Name = "contacts_" + jobinfo.Directory.Name };
                var nameColumn = contactsTable.Columns.Add("name", typeof(string));
                var typeColumn = contactsTable.Columns.Add("type", typeof(string));
                var keyColumn = contactsTable.Columns.Add("key", typeof(string));
                var valueColumn = contactsTable.Columns.Add("value", typeof(string));
                contactsTable.PrimaryKeys = new Table.Column[] { keyColumn };
                foreach (var kvp in contacts)
                    contactsTable.Rows.Add(new string[] { jobinfo.Name, documentTypeLanguage.ToString(), kvp.Key, kvp.Value });
                return Tuple.Create(itemTable, documentTypeLanguage, contactsTable);
            }
        }
        private static void SortWords(List<Word> words)
        {
            words.Sort((w1, w2) =>
            {
                int lvl1 = LineNbr_round(w2.BoundingBox.Bottom).CompareTo(LineNbr_round(w1.BoundingBox.Bottom));
                if (lvl1 != 0) return lvl1;
                int lvl2 = w1.BoundingBox.Left.CompareTo(w2.BoundingBox.Left);
                return lvl2;
            });
        }
        private const byte roundFactor = 1;
        private static double LineNbr_round(double Y)=>  Math.Round(Math.Round(Y / roundFactor, 0) * roundFactor);
        internal static Dictionary<string, string> Get_contactKeysValues(Dictionary<double, Dictionary<byte, Dictionary<int, Letter>>> lines, double col3_qtyLeft)
        {
            /// still get --> " eE)-Mail:   nicole.segal@flaghouse.com"
            /// since the right parenthesis of "(Offre spéciale fin d'année) is overlapped with "E-Mail:" 
            var keysAndValues = new Dictionary<string, string>();
            var leftLines = new List<string>();
            var rightLines = new List<string>();
            foreach (var line in lines.Values)
            {
                var leftWords = new List<string>();
                var rightWords = new List<string>();
                foreach (var word in line.Values)
                {
                    var letters = new List<Letter>(word.Values);
                    letters.Sort((l1, l2) => l1.StartBaseLine.X.CompareTo(l2.StartBaseLine.X));
                    var stringOfLetters = string.Join("", letters.Where(l => l.EndBaseLine.X < col3_qtyLeft).Select(l => l.Value));
                    leftWords.Add(stringOfLetters);
                    stringOfLetters = string.Join("", letters.Where(l => l.EndBaseLine.X >= col3_qtyLeft).Select(l => l.Value));
                    rightWords.Add(stringOfLetters);
                }
                var stringOfWords = string.Join(" ", leftWords);
                leftLines.Add(stringOfWords);
                stringOfWords = string.Join(" ", rightWords);
                rightLines.Add(stringOfWords);
            }
            return keysAndValues;
        }
        private static Dictionary<string, string> Get_contactKeysValues(List<Word> words, double col3_qtyLeft, byte margin)
        {
            var keysAndValues = new Dictionary<string, string>();
            SortWords(words);
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
                //if (leftSide.ContainsValue("NSOPP6009")) Debugger.Break();
                foreach (var newKeyValue in leftSide)
                    if (newKeyValue.Key.StartsWith("Address"))
                    {
                        int addressLineCount = keysAndValues.Keys.Where(k => k.StartsWith("Address")).Count();
                        keysAndValues.Add($"Address-{1 + addressLineCount}", newKeyValue.Value);
                    }
                    else if (keysAndValues.ContainsKey(newKeyValue.Key))
                    {
                        string lastValue = keysAndValues[newKeyValue.Key].Trim();
                        string newValue = newKeyValue.Value.Trim();
                        var values = new List<string> { lastValue, newValue };
                        values.Sort((v1, v2) => v2.Length.CompareTo(v1.Length));
                        keysAndValues[newKeyValue.Key] = values.First();
                    }
                    else if (!keysAndValues.ContainsKey(newKeyValue.Key))
                        keysAndValues.Add(newKeyValue.Key, newKeyValue.Value);

                var rightSide = Get_bySide(words_AboveBelow, col3_qtyLeft, false);
                foreach (var newKeyValue in rightSide)
                    if (newKeyValue.Key.StartsWith("email"))
                    {
                        string emailKey = "email" + (newKeyValue.Value.ToLowerInvariant().Contains("flaghouse") ? "-rep" : "-client");
                        keysAndValues[emailKey] = newKeyValue.Value;
                    }
                    else if (!keysAndValues.ContainsKey(newKeyValue.Key))
                        keysAndValues.Add(newKeyValue.Key, newKeyValue.Value);
            }
            return keysAndValues;
        }
        private static Dictionary<string, string> Get_bySide(List<Word> words, double col3_qtyLeft, bool isLeftSide = true)
        {
            var keysAndValues = new Dictionary<string, string>();
            var words_OneSide = new List<Word>();
            var text = string.Join(" ", words.Select(w => w.Text));

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
                string wordKey = string.Join(" ", words_colonLeft.Select(w => w.Text));
                // a pdfword is a string of letters without spaces, so sometimes the value is not a separate word --> QUOTE#:NSOPP5357
                string wordValue = string.Empty;
                string[] keyValue = wordKey.Split(':');
                if (keyValue.Length > 1)
                {
                    wordKey = keyValue[0];
                    wordValue = keyValue[1].Trim();
                }
                //if(text.Contains("E-Mail")) Debugger.Break();
                wordValue += words_colonRight.Any() ? string.Join(" ", words_colonRight.Select(w => w.Text)) : string.Empty;
                string normalizedKey = Get_key(wordKey); 
                if (!keysAndValues.ContainsKey(normalizedKey)) keysAndValues[normalizedKey] = wordValue;
                //if (words.Where(w => w.Text.Contains("QUOTE")).Any() & words_colonLeft.Where(w => w.Text.Contains("QUOTE")).Any()) Debugger.Break();
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
                    if (Regex.IsMatch(wrappedText, "[A-Z]{3}PP[0-9]{4}"))
                        keysAndValues.Add("Quote#", wrappedText);

                    // MONTREAL, QC H3W 1J6 CANADA
                    else if (Regex.IsMatch(wrappedText, "[ABCEGHJ-NPRSTVXY]\\d[ABCEGHJ-NPRSTV-Z][ -]?\\d[ABCEGHJ-NPRSTV-Z]\\d"))
                        foreach (var cpp in CityProvincPostal(wrappedText))
                            keysAndValues.Add(cpp.Key, cpp.Value);

                    // 6333 RUE DE TERREBONNE
                    else if (isLeftSide) // no colon, left side -> address line
                        keysAndValues.Add("Address", wrappedText);
                }
            }
            if (keysAndValues.ContainsKey("Fax#")) keysAndValues.Remove("Fax#"); // nobody uses a fax machine
            return keysAndValues;
        }
        private static string Get_key(string keyIn)
        {
            if (keyIn == null) return string.Empty;
            if (Regex.IsMatch(keyIn, "date", RegexOptions.IgnoreCase)) return "Date";
            if (Regex.IsMatch(keyIn, "(client|customer|Compagnie)", RegexOptions.IgnoreCase)) return "Customer";
            if (Regex.IsMatch(keyIn, "(order|commande|r(e|é)f(e|é)rence)", RegexOptions.IgnoreCase)) return "Order#";
            if (Regex.IsMatch(keyIn, "(contact|attention)", RegexOptions.IgnoreCase)) return "Contact";
            if (Regex.IsMatch(keyIn, "(t(e|é)l(e|é)){0,1}phone {0,}#{0,1}", RegexOptions.IgnoreCase)) return "Phone#";
            if (Regex.IsMatch(keyIn, "(e-{0,1}mail|courriel|-mail)", RegexOptions.IgnoreCase)) return "email";
            if (Regex.IsMatch(keyIn, "(consultante{0,1}|((fh|flaghouse) ){0,1}rep)", RegexOptions.IgnoreCase)) return "Rep";
            if (Regex.IsMatch(keyIn, "(adresse|address)", RegexOptions.IgnoreCase)) return "Address";
            if (Regex.IsMatch(keyIn, "(quote|soumission) {0,}#", RegexOptions.IgnoreCase)) return "Quote#"; // not necessary as is handled in Get_bySide- !words_colon.Any()
            if (Regex.IsMatch(keyIn, "(t(e|é)l(e|é)copieur|fax|facsimile)", RegexOptions.IgnoreCase)) return "Fax#"; // who tf still uses a fax machine???
            return keyIn;
        }
        private static Dictionary<string, string> CityProvincPostal(string address)
        {
            if (address == null) return null;
            var keysAndValues = new Dictionary<string, string>
            {
                { "City", string.Empty },
                { "Province", string.Empty },
                { "PostalCode", string.Empty }
            };
            var postalMatch = Regex.Match(address, "[ABCEGHJ-NPRSTVXY]\\d[ABCEGHJ-NPRSTV-Z][ -]?\\d[ABCEGHJ-NPRSTV-Z]\\d", RegexOptions.IgnoreCase);
            if (postalMatch.Success)
            {
                string postalCode = postalMatch.Value.Replace(" ", string.Empty).ToUpperInvariant();
                keysAndValues["PostalCode"] = $"{postalCode.Substring(0, 3)} {postalCode.Substring(3, 3)}".ToUpperInvariant();
                var cityProvince = address.Substring(0, postalMatch.Index);
                var addressElements = cityProvince.Split(',');
                keysAndValues["City"] = string.Join(", ", addressElements.Take(addressElements.Length - 1).Select(c => ProperCase(c)));
                var province = addressElements.Last().Trim();
                var provinceFullname = string.Empty;
                bool foundProvince = false;
                // english- fullname
                foreach (var provPattern in provinces)
                {
                    if (Regex.IsMatch(province, provPattern.Item1, RegexOptions.IgnoreCase))
                    {
                        provinceFullname = provPattern.Item1;
                        foundProvince = true;
                        break;
                    }
                }
                if (!foundProvince)
                {
                    // french- fullname
                    foreach (var provPattern in provinces)
                    {
                        if (Regex.IsMatch(province, provPattern.Item2, RegexOptions.IgnoreCase))
                        {
                            provinceFullname = provPattern.Item1;
                            foundProvince = true;
                            break;
                        }
                    }
                }
                if (!foundProvince)
                {
                    // standard abbreviations
                    foreach (var provPattern in provinces)
                    {
                        if (Regex.IsMatch(province, provPattern.Item3, RegexOptions.IgnoreCase))
                        {
                            provinceFullname = provPattern.Item1;
                            foundProvince = true;
                            break;
                        }
                    }
                }
                keysAndValues["Province"] = foundProvince ? provinceFullname : province;
                //if (address.ToLower().Contains("monkey")) Debugger.Break();
            }
            return keysAndValues;
        }
        private static string ProperCase(string phraseIn)
        {
            if (phraseIn == null) return null;
            var propercase = new List<string>();
            var words = phraseIn.Trim().Split(' ');
            foreach (var word in words)
            {
                if (Regex.Split(word, "[^a-zâêîôûàèìòùëïüçé]{1,}", RegexOptions.IgnoreCase).Length == 1)
                    propercase.Add($"{word.Substring(0, 1).ToUpperInvariant()}{word.Substring(1, word.Length - 1).ToLowerInvariant()}");
                else
                {
                    var letters = new Dictionary<bool, Dictionary<int, List<char>>>
                    {
                        {true, new Dictionary<int, List<char>>() },
                        {false, new Dictionary<int, List<char>>() }
                    };
                    bool lastWasLetter = false;
                    for (int l = 0; l < word.Length; l++)
                    {
                        char letter = word[l];
                        bool isLetter = Regex.IsMatch(letter.ToString(), "[a-z]", RegexOptions.IgnoreCase);
                        if (l == 0 | lastWasLetter != isLetter)
                        {
                            // first letter is series, create collection
                            letters[isLetter][letters[isLetter].Count] = new List<char>();
                            lastWasLetter = isLetter;
                        }
                        letters[isLetter][letters[isLetter].Count - 1].Add(letter);
                    }
                    // true  [0]l [1]original
                    // false [0]'
                    var maxLetterGrp = new int[] { letters[true].Count, letters[false].Count }.Max();
                    var properString = string.Empty;
                    for (var i = 0; i < maxLetterGrp; i++)
                        foreach (var isLetter in new bool[] { true, false })
                            if (letters[isLetter].ContainsKey(i))
                            {
                                var wordSegment = string.Join("", letters[isLetter][i]);
                                properString += isLetter ? ProperCase(wordSegment) : wordSegment;
                            }
                    propercase.Add(properString);
                }
            }
            return string.Join(" ", propercase);
        }
        private static readonly List<Tuple<string, string, string, string>> provinces = new List<Tuple<string, string, string, string>>
        {
            Tuple.Create("Newfoundland and Labrador","Terre-Neuve-et-Labrador","NL","N\\.{0,1}L\\.{0,1}|T\\.{0,1}-N\\.{0,1}-L\\.{0,1}"),
            Tuple.Create("Prince Edward Island","Île-du-Prince-Édouard","PE","(P\\.{0,1}E\\.{0,1}I\\.{0,1}|Î\\.{0,1}-P\\.{0,1}-É\\.{0,1})"),
            Tuple.Create("Nova Scotia","Nouvelle-Écosse","NS","(N\\.{0,1}S\\.{0,1}|N\\.{0,1}-É\\.{0,1})"),
            Tuple.Create("New Brunswick","Nouveau-Brunswick","NB","(N\\.{0,1}B\\.{0,1}|N\\.{0,1}-B\\.{0,1})"),
            Tuple.Create("Québec","Qu(e|é)bec","QC","(Que\\.{0,1}|Qc)"),
            Tuple.Create("Ontario","Ontario","ON","Ont\\.{0,1}"),
            Tuple.Create("Manitoba","Manitoba","MB","(Man\\.{0,1}|Man\\.{0,1})"),
            Tuple.Create("Saskatchewan","Saskatchewan","SK","Sask\\.{0,1}"),
            Tuple.Create("Alberta","Alberta","AB","(Alta\\.{0,1}|Alb\\.{0,1})"),
            Tuple.Create("British Columbia","Colombie-Britannique","BC","(B\\.{0,1}C\\.{0,1}|C\\.{0,1}-B\\.{0,1})"),
            Tuple.Create("Yukon","Yukon","YK","(Y\\.{0,1}T\\.{0,1}|Yn)"),
            Tuple.Create("Northwest Territories","Territoires du Nord-Ouest","NT","(N\\.{0,1}W\\.{0,1}T\\.{0,1}|T\\.{0,1}N\\.{0,1}-O\\.{0,1})"),
            Tuple.Create("Nunavut","Nunavut","NT","(Nvt\\.{0,1}|Nt)")
        };
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
        internal static string Printout()
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