using System;
using System.Text;
using Soneta.Business;
using System.Collections;
using Soneta.Handel;
using Soneta.Kasa;
using Soneta.Types;
using Soneta.Langs;

[assembly: Worker(typeof(enovaExt.TenderHut_ExportSprzedaz_MULTI), typeof(DokumentHandlowy))]

namespace enovaExt
{
    public class TenderHut_ExportSprzedaz_MULTI
    {
        //[Context]
        public DokHandlowe dokHandlowe { get; set; }

        public Boolean bylo;
        Date from = Date.Empty;
        Date to = Date.Empty;

        [Context]
        public Context Context { get; set; }

        [Action(
            "Eksport - Sprzedaż",
            Priority = 2,
            Icon = ActionIcon.Wizard,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        [TranslateIgnore]
        public NamedStream SaveDataWorker()
        {
            from = ((FromTo)Context[typeof(FromTo)]).From;
            to = ((FromTo)Context[typeof(FromTo)]).To;

            string fileName = Context.Login.Database.Name + "_" + from.ToString().Substring(0, 7) + ".csv";
            /*
            return new NamedStream(fileName,
                () =>
                {
                    var writter = new System.IO.StreamWriter(new System.IO.MemoryStream(), System.Text.Encoding.UTF8);
                    WriteHeader(writter);
                    GetDataToSave(writter);
                    writter.Flush();
                    //plik.Close();
                    return ((System.IO.MemoryStream)writter.BaseStream).ToArray();
                });
            */
            var writter = new System.IO.StreamWriter(new System.IO.MemoryStream(), System.Text.Encoding.UTF8);
            WriteHeader(writter);
            GetDataToSave(writter);
            writter.Flush();

            return new NamedStream(fileName, ((System.IO.MemoryStream)writter.BaseStream).ToArray());
        }

        public void GetDataToSave(System.IO.StreamWriter plik)
        {
            //INavigatorContext nav = Context[typeof(INavigatorContext)] as INavigatorContext;
            View nav = Context[typeof(View)] as View;
            string[] array = new string[32];
            HandelModule hm = HandelModule.GetInstance(Context.Session);
            Date okres = ((FromTo)Context[typeof(FromTo)]).From;
            //View dh = hm.DokHandlowe.CreateView();

            //dh.Condition &= new FieldCondition.GreaterEqual("Data", from);
            //dh.Condition &= new FieldCondition.LessEqual("Data", to);
            //dh.Condition &= new FieldCondition.LessEqual("Kategoria", KategoriaHandlowa.Sprzedaż);

            //foreach (DokumentHandlowy dok in dh)
            //{

            decimal factor = Decimal.Zero;

            foreach (DokumentHandlowy dok in nav)
            {
                if (dok.Definicja.Symbol == "FPRO" || dok.Definicja.Symbol == "FZAL")
                    continue;

                foreach (PozycjaDokHandlowego pozDok in dok.Pozycje)
                {
                    //                dh = (DokumentHandlowy)o;
                    Hashtable ht = new Hashtable();
                    Hashtable htworker = new Hashtable();

                    array[0] = dok.Session.Global.Features["SELLER"].ToString(); //dok.Wydruk.Naglowek.Dane.Dane.Nazwa;
                    array[1] = dok.NumerPelnyZapisany;
                    array[2] = dok.Kontrahent.Nazwa;
                    array[3] = dok.Dostawa.Termin.ToString(); //  dok.DataOperacji.ToString();
                    array[4] = dok.Platnosci.Count > 0 ? dok.Platnosci.GetFirst().Termin.ToString() : "";

                    if (dok.Korekta)
                    {
                        array[5] = (pozDok.Suma.BruttoCy.Value - pozDok.PozycjaKorygowana.Suma.BruttoCy.Value).ToString(); //   .Suma.BruttoCy.Value.ToString(); //   dok.Suma.BruttoCy.Value.ToString();
                        array[8] = (pozDok.Suma.Brutto - pozDok.PozycjaKorygowana.Suma.Brutto).ToString();
                        factor = Math.Abs((pozDok.Suma.Brutto - pozDok.PozycjaKorygowana.Suma.Brutto) / dok.Suma.Brutto);
                    }
                    else
                    {
                        array[5] = pozDok.Suma.BruttoCy.Value.ToString(); //   dok.Suma.BruttoCy.Value.ToString();
                        array[8] = pozDok.Suma.Brutto.ToString();  //I
                        factor = Math.Abs(pozDok.Suma.Brutto / dok.Suma.Brutto);
                    }

                    array[6] = pozDok.Suma.BruttoCy.Symbol;//  WalutaKontrahenta.ToString();
                    array[7] = dok.KursWaluty.ToString();
                    //array[8] = pozDok.Suma.Brutto.ToString();  //I

                    StanRozliczeniaRozrachunkuWorker worker = new StanRozliczeniaRozrachunkuWorker();
                    //worker.StanRozliczenia = StanRozliczeniaRozrachunku.Nierozliczone;
                    decimal kwota = Decimal.Zero;
                    decimal rozliczono = Decimal.Zero;
                    decimal niezaplacone = Decimal.Zero;
                    decimal czesciowo = Decimal.Zero;
                    decimal resztaRozl = Decimal.Zero;

                    Date dataRozliczenia = Date.MinValue;
                    int przeterm = 0;
                    //Date dataplatnosci = Date.MinValue;

                    foreach (Platnosc pl in dok.Platnosci)
                        foreach (RozrachunekIdx idx in pl.Rozrachunki)
                        {
                            worker.RozrachunekIdx = idx;
                            if (worker.Kwota.Symbol == "PLN")
                                kwota = pl.Kwota.Value; //worker.Kwota.Value;
                            else
                                kwota = pl.KwotaKsiegi.Value;// * (decimal)dok.KursWaluty;

                            rozliczono += worker.KwotaRozliczona.Value * (decimal)dok.KursWaluty;
                            niezaplacone += kwota - rozliczono;
                            czesciowo += rozliczono;

                            if (idx.DataRozliczenia > dataRozliczenia)
                                dataRozliczenia = idx.DataRozliczenia;

                            if (przeterm < idx.ZwłokaNaDzień(Date.Now)) // idx.PrzeterminowanoDni)
                                przeterm = idx.ZwłokaNaDzień(Date.Now);
                        }

                    //rozliczono = rozliczono * (decimal)dok.KursWaluty;
                    //niezaplacone = niezaplacone * (decimal)dok.KursWaluty;

                    array[9] = (factor * rozliczono).ToString(); //J

                    if (dok.Korekta)
                        if (pozDok.Suma.Brutto - pozDok.PozycjaKorygowana.Suma.Brutto < 0)
                            array[9] = (-1 * factor * rozliczono).ToString(); //J

                    array[10] = (factor * niezaplacone).ToString(); //K
                    array[11] = dataRozliczenia != Date.MaxValue ? dataRozliczenia.ToString() : ""; // L  - Real payment date   
                    array[12] = dok.Platnosci.Count > 0 ? dok.Platnosci.GetFirst().Termin.ToString() : ""; //dok.Platnosci.GetFirst().Termin.ToString();
                    array[13] = Math.Abs(dok.Suma.Brutto) <= rozliczono ? "YES" : "NO";
                    array[14] = Date.Now.ToString(); // O   - today date
                    array[15] = przeterm.ToString(); // "=O2-E2";  //przeterm.ToString();
                    array[16] = "";
                    array[17] = "";

                    if (dok.Korekta)
                    {
                        array[18] = (pozDok.Suma.Netto - pozDok.PozycjaKorygowana.Suma.Netto).ToString();
                    }
                    else
                        array[18] = pozDok.Suma.Netto.ToString();   //dok.Suma.Netto.ToString(); //   "PLN";


                    if (array[11] != "")
                        array[19] = ((dok.Platnosci.Count > 0 ? dok.Platnosci.GetFirst().Termin : Date.Empty) - (dataRozliczenia != Date.MaxValue ? dataRozliczenia : Date.Empty)).ToString();
                    else
                        array[19] = (Date.Today - (dok.Platnosci.Count > 0 ? dok.Platnosci.GetFirst().Termin : Date.Empty)).ToString();

                    //array[19] = "=JEŻELI(L2>0;E2-L2;JEŻELI(L2=\"\";O2-E2))";  // T  - actual overdue days
                    array[20] = dok.Kontrahent.Adres.Kraj;


                    if (pozDok.Features["Sales_responsible"] != null)
                        array[21] = pozDok.Features["Sales_responsible"].ToString();
                    else
                        array[21] = "";

                    //if (dok.Features["Partner"] != null)
                    //  array[21] = dok.Features["Partner"].ToString();
                    //else
                    //  array[21] = "";

                    array[23] = okres.Month.ToString();
                    array[22] = okres.Year.ToString();
                    /*
                    if (dok.Ewidencja.DataEwidencji != Date.Empty)
                        array[24] = dok.Ewidencja.DataEwidencji.ToString();
                    else
                        array[24] = "";
                    */
                    array[24] = dok.Features["Data zatwierdzenia faktury"].ToString();

                    if (pozDok.Towar.GrupaTowarowaVat != null)
                        array[25] = pozDok.Towar.GrupaTowarowaVat.Symbol;
                    else
                        array[25] = "";


                    Soneta.Ksiega.ElemSlownika elem = null;// (Soneta.Ksiega.ElemSlownika)pozDok.Features["Projekt"];

                    if (pozDok.Features["Projekt"] != null)
                    {
                        elem = (Soneta.Ksiega.ElemSlownika)pozDok.Features["Projekt"];

                        //if (!array[17].Contains(pozDok.Features["Projekt"].ToString()))
                        array[17] = elem.Nazwa + " (" + elem.Symbol + ")"; // pozDok.Features["Projekt"].ToString() + " ";
                    }

                    string industry = dok.Kontrahent.Features["Industries"].ToString();
                    industry = industry + " (" + industry.Substring(0, industry.IndexOf(" ")) + ")";

                    array[26] = industry; // dok.Kontrahent.Features["Industries"].ToString();

                    array[27] = array[0];

                    if (dok.Korekta)
                        array[28] = (pozDok.Suma.NettoCy.Value - pozDok.PozycjaKorygowana.Suma.NettoCy.Value).ToString();
                    else
                        array[28] = pozDok.Suma.NettoCy.Value.ToString();

                    array[29] = pozDok.Suma.NettoCy.Symbol;

                    if (pozDok.Features["Segment"] != null)
                        array[30] = pozDok.Features["Segment"].ToString();
                    else
                        array[30] = "";

                    if (pozDok.Features["Product_group"] != null)
                    {
                        elem = (Soneta.Ksiega.ElemSlownika)pozDok.Features["Product_group"];
                        array[31] = elem.Nazwa + " (" + elem.Symbol + ")"; //pozDok.Features["Product_group"].ToString();
                    }
                    else
                        array[31] = "";

                    StringBuilder builder = new StringBuilder();

                    foreach (string value in array)
                    {
                        builder.Append(value);
                        builder.Append(';');
                    }

                    plik.WriteLine(builder.ToString());
                }
            }
        }

        private void WriteHeader(System.IO.StreamWriter writter)
        {
            //writter.WriteLine("Seller;Invoice No;Buyer;Sales date;Due date;Gross Value;Currency;Exchange rate;PLN Gross Value;Payment value;To be paid;Real payment date;Expected payment date;" +
            //"Paid;today date;overdue by days;Reminder mail;Contains / PROJEKT;PLN;actual overdue days;Country;Partner;Y;M;Invoice sent;GTU;Industry;com;currency net value;Currency2;Our reporting segment;Product group");

            writter.WriteLine("Seller;Invoice No;Buyer;Sales date;Due date;Gross Value;Currency;Exchange rate;PLN Gross Value;Payment value;To be paid;Real payment date;Expected payment date;" +
                        "Paid;today date;overdue by days;Reminder mail;Contains / PROJEKT;PLN;actual overdue days;Country;Partner;Y;M;Invoice sent;GTU;Industry;com;currency net value;Currency2;Our reporting segment;Product group");
        }

        public static bool IsVisibleFunkcja(DokumentHandlowy dokument)
        {
            return dokument.Definicja.Kategoria == KategoriaHandlowa.Sprzedaż || dokument.Definicja.Kategoria == KategoriaHandlowa.KorektaSprzedaży;
            //return true;
        }
    }
}

