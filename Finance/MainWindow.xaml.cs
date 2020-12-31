using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;
using System.Text;

namespace Finance
{
    /// <summary>
    /// Interakční logika pro MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Random rnd = new Random();
        OleDbConnection connection;
        OleDbCommand command;
        OleDbDataAdapter dataAdapter;
        OleDbDataReader reader;
        private string zvolenaDB, zvolenaDBZobr;
        private string defaultDB, defaultZobrNazevDB, cestaDB;

        private string myConfigDefaults = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><myConfiguration><programInfo><lastStateChange>" + DateTime.Now + "</lastStateChange></programInfo><defaultDB>Hotovost</defaultDB><defaultZobrNazevDB>Hotovost</defaultZobrNazevDB><pathDB>Finance.accdb</pathDB><dbBackups><pathBackups>/backups/</pathBackups><backupsAmount>5</backupsAmount></dbBackups></myConfiguration>";

        private string cestaBackups;
        private int pocetUchovavanychZaloh;
        
        private string datumVstup;
        private double castkaVstup;
        private string poznamkaVstup;



        public MainWindow()
        {
            InitializeComponent();
            
        }
        private bool NacistDefaultniNastaveni()
        {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load("myConfig.xml");

                defaultDB = doc.DocumentElement.SelectSingleNode("/myConfiguration/defaultDB").InnerText;
                defaultZobrNazevDB = doc.DocumentElement.SelectSingleNode("/myConfiguration/defaultZobrNazevDB").InnerText;
                cestaDB = doc.DocumentElement.SelectSingleNode("/myConfiguration/pathDB").InnerText;
                cestaBackups = doc.DocumentElement.SelectSingleNode("/myConfiguration/dbBackups/pathBackups").InnerText;
                pocetUchovavanychZaloh = Convert.ToInt32(doc.DocumentElement.SelectSingleNode("/myConfiguration/dbBackups/backupsAmount").InnerText);

                doc.DocumentElement.SelectSingleNode("/myConfiguration/programInfo/lastStateChange").InnerText = DateTime.Now.ToString();


                doc.Save("myConfig.xml");
                return true;
            }
            catch (FileNotFoundException)
            {
                MessageBoxResult result = MessageBox.Show("Nepodařilo se nalézt konfigurační soubor. Pokud chcete vygenerovat nový konfigurační soubor a pokračovat ve spuštění programu, klikněte na OK. V opačném případě bude program ukončen.", "Chyba", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    if (!File.Exists("myConfig.xml")) File.WriteAllText("myConfig.xml", myConfigDefaults);
                    NacistDefaultniNastaveni();
                    return false;
                }
                else
                {
                    Close();
                    return false;
                }
            }
            catch (NullReferenceException)
            {
                //VS doesn't catch it in Debug or Release mode
                MessageBoxResult result = MessageBox.Show("Konfigurační soubor je pravděpodobě poškozený. Pokud chcete vygenerovat nový konfigurační soubor a pokračovat ve spuštění programu, klikněte na OK. V opačném případě bude program ukončen.", "Chyba", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    File.WriteAllText("myConfig.xml", myConfigDefaults);
                    NacistDefaultniNastaveni();
                    return false;
                }
                else
                {
                    Close();
                    return false;
                }
            }
            catch (XmlException)
            {
                MessageBoxResult result = MessageBox.Show("Konfigurační soubor je pravděpodobě poškozený. Pokud chcete vygenerovat nový konfigurační soubor a pokračovat ve spuštění programu, klikněte na OK. V opačném případě bude program ukončen.", "Chyba", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    File.WriteAllText("myConfig.xml", myConfigDefaults);
                    NacistDefaultniNastaveni();
                    return false;
                }
                else
                {
                    Close();
                    return false;
                }
            }
            catch (Exception e)
            {
                string innerEx = e.ToString();
                MessageBox.Show($"Vyskytla se neočekávaná chyba, kvůli které se program ukončí." +
                    $"\n\rInner Exception: {innerEx} " +
                    $"\n\rException Message: {e.Message}", "Fatální chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                Close();
                return false;
            }
            finally
            {
                doc = null;
            }
        }

        private void PoLoginu()
        {
            if (NacistDefaultniNastaveni())
            {
                zvolenaDB = defaultDB;

                DPDatum.SelectedDate = DateTime.Now;

                command = new OleDbCommand();
                AktualizujStavCelkovehoStavuFinanci();
                NaplnTabulkuDatyZeZvoleneDT(defaultDB, defaultZobrNazevDB);

            }
        }

        private void ZalohovaniDB()
        {
            //todo: nastavit jako vyskakovací okno
            //todo: zavest checkbox na login window s automatickou zálohou
            
            //kontrola existence složky
            if (!Directory.Exists(cestaBackups)) Directory.CreateDirectory(cestaBackups);

            //načtení záloh ze složky (datum)
            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory() + @"\backups", "*.accdb");
            List<KeyValuePair<DateTime, string>> filePairs = new List<KeyValuePair<DateTime, string>>();
            foreach (var item in files) filePairs.Add(new KeyValuePair<DateTime, string>(File.GetLastWriteTime(item), item));

            //řazení souborů podle data změny
            var serazene = filePairs.OrderBy(x => x.Key).ToList();
            DateTime thisDB = File.GetLastWriteTime(cestaDB);

            //pokud se najde stejný soubor se stejným datem změny, metoda se ukončí
            foreach (var item in serazene) if (item.Key == thisDB) { return; }

            //generování náhodného konce názvu souboru 
            string randomEnd = string.Empty;
            for (int i = 0; i < 5; i++)
            {
                if (Convert.ToBoolean(rnd.Next(0, 2)))
                {
                    var x = (char)rnd.Next(65, 91);
                    randomEnd += x;
                }
                else
                {
                    var x = (char)rnd.Next(48, 58);
                    randomEnd += x;
                }
            }
            
            //mazání přebytečných záloh
            if (filePairs.Count > pocetUchovavanychZaloh)
            {    
                int pocetSoucasnychZaloh = serazene.Count;
                int iterace = ((pocetUchovavanychZaloh - pocetSoucasnychZaloh) * -1) +1;
                for (int i = 0; i < iterace; i++)
                {
                    var value = serazene[0];
                    File.Delete(value.Value);
                    serazene.Remove(value);
                }
            }

            //vytváření zálohy
            File.Copy(cestaDB, cestaBackups + $@"\{DateTime.Now.ToString().Substring(0,10)}-{randomEnd}.accdb");
        }

        private void AktualizujStavCelkovehoStavuFinanci()
        {
            List<double> list = new List<double>();
            list.AddRange(VratCastkyZeZvoleneDB("Hotovost"));
            list.AddRange(VratCastkyZeZvoleneDB("BankaBeznyUcet"));
            list.AddRange(VratCastkyZeZvoleneDB("BankaSporiciUcet"));
            double soucet = list.Sum();
            LBCelkovyZustatek.Content = String.Format("{0} Kč", soucet);
            AktualizujAktZustatekZvoleneDB(zvolenaDB);
        }

        private void AktualizujAktZustatekZvoleneDB(string db)
        {
            LBAktZustatekDT.Content = VratCastkyZeZvoleneDB(db).Sum().ToString() + " Kč";
        }

        //todo: přepsat na async
        private List<double> VratCastkyZeZvoleneDB(string dbName)
        {
            List<double> list = new List<double>();
            if (connection.State != ConnectionState.Open) connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT Castka from " + dbName);
            command.Connection = connection;

            reader = command.ExecuteReader();
            while (reader.Read())
            {
                list.Add(Convert.ToDouble(reader[0].ToString()));
            }
            if (connection.State != ConnectionState.Closed) connection.Close();
            return list;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi)
            {
                switch (mi.Name.ToString())
                {
                    case "MIProgram_Odhlasit":
                        {
                            OdhlaseniUzivatele();
                            ZalohovaniDB();
                            break;
                        }
                    case "MIProgram_Ukoncit":
                        {
                            OdhlaseniUzivatele();
                            Close();
                            break;
                        }
                    case "MIDatabaze_Hotovost":
                        {
                            NaplnTabulkuDatyZeZvoleneDT("Hotovost", "Hotovost");
                            break;
                        }
                    case "MIDatabaze_BankaBeznyUcet":
                        {
                            NaplnTabulkuDatyZeZvoleneDT("BankaBeznyUcet", "Běžný účet");
                            break;
                        }
                    case "MIDatabaze_BankaSporiciUcet":
                        {
                            NaplnTabulkuDatyZeZvoleneDT("BankaSporiciUcet", "Spořící účet");
                            break;
                        }
                    default: break;
                }
            }
        }

        private void NaplnTabulkuDatyZeZvoleneDT(string nazevDT, string zobrVybranaDT)
        {
            zvolenaDB = nazevDT;
            zvolenaDBZobr = zobrVybranaDT;
            LBZvolenaDT.Content = zobrVybranaDT;

            if (connection.State != ConnectionState.Open) connection.Open();
            command = new OleDbCommand("SELECT ID, Datum, Castka as Částka, Poznamka as Poznámka from " + nazevDT + " order by ID desc");
            command.Connection = connection;

            dataAdapter = new OleDbDataAdapter(command);
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);
            DGrDatabaze.ItemsSource = dt.AsDataView();
            AktualizujAktZustatekZvoleneDB(nazevDT);
            if (connection.State != ConnectionState.Closed) connection.Close();


        }


        private bool ValidaceVstupu(object datum, string castka, string poznamka)
        {
            try
            {
                DateTime convertedDatum;
                if (datum == null) convertedDatum = DateTime.Now;
                else convertedDatum = (DateTime)DPDatum.SelectedDate;
                datumVstup = convertedDatum.ToString();

                castkaVstup = Convert.ToDouble(castka);

                poznamkaVstup = System.Text.RegularExpressions.Regex.Replace(poznamka, @"[^a-z^0-9^ ^-^ ^(^ ^)^ _ ^Á^ ^Č^ ^Ď^ ^É^ ^Ě^ ^Í^ ^Ň^ ^Ř^ ^Š^ ^Ť^ ^Ů^ ^Ú^ ^Ý^ ^Ž^ ^á^ ^č^ ^ď^ ^é^ ^ě^ ^í^ ^ň^ ^ó^ ^ř^ ^š^ ^ť^ ^ú^ ^ů^ ^ý^ ^ž^ ^/^ ^+^ ^,^ ^.^ ^\-^]", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (castkaVstup == 0 || poznamkaVstup.Length <= 0)
                {
                    return false;
                }

                return true;
            }
            catch (Exception e)
            {
                string innerEx = e.InnerException == null ? "-" : e.InnerException.ToString();
                MessageBox.Show($"Vyskytla se neočekávaná chyba. Nebyly provedené žádné změny." +
                    $"\n\rInner Exception: {innerEx} " +
                    $"\n\rException Message: {e.Message}", "Fatální chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private void BTPridatZaznam_Click(object sender, RoutedEventArgs e)
        {
            if (DGrDatabaze.SelectedItems.Count == 1)
            {
                DPDatum.SelectedDate = DateTime.Now;
            }
            else if (DGrDatabaze.SelectedItems.Count > 1)
            {
                MessageBox.Show("Bylo vybráno více záznamů ke kopírování, což nelze provést, proto byl zkopírován pouze první vybraný záznam.", "Informace", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            if (ValidaceVstupu(DPDatum.SelectedDate, TxBxCastka.Text, TxBxPoznamka.Text))
            {
                command = new OleDbCommand();
                if (connection.State != ConnectionState.Open) connection.Open();
                command.Connection = connection;

                command.CommandText = "insert into " + zvolenaDB +
                    " (Datum,Castka,Poznamka) " +
                    "Values(@datum,@castka,@poznamka)";
                command.Parameters.AddWithValue("@datum", datumVstup.ToString());
                command.Parameters.AddWithValue("@castka", castkaVstup);
                command.Parameters.AddWithValue("@poznamka", poznamkaVstup);
                command.ExecuteNonQuery();
                NaplnTabulkuDatyZeZvoleneDT(zvolenaDB, zvolenaDBZobr);
                VymazDataZFormu();
                AktualizujStavCelkovehoStavuFinanci();
                if (connection.State != ConnectionState.Closed) connection.Close();
            }
            else MessageBox.Show("Zadané vstupní hodnoty neprošly validací. Zkuste hodnoty zadat znovu.", "Chyba", MessageBoxButton.OK, MessageBoxImage.Error);
            BTUpravitZaznam.Visibility = Visibility.Hidden;
        }

        private void VymazDataZFormu()
        {
            DPDatum.SelectedDate = DateTime.Now;
            TxBxCastka.Text = "0";
            TxBxPoznamka.Text = string.Empty;
        }

        private void BTUpravitZaznam_Click(object sender, RoutedEventArgs e)
        {
            if (DGrDatabaze.SelectedItems.Count > 0)
            {
                if (ValidaceVstupu(DPDatum.SelectedDate, TxBxCastka.Text, TxBxPoznamka.Text))
                {
                    if (DGrDatabaze.SelectedItems.Count > 1)
                    {
                        MessageBox.Show("Bylo vybráno více záznamů k úpravě, což nelze provést, proto byl upraven pouze první vybraný záznam.", "Informace", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    DataRowView radek = (DataRowView)DGrDatabaze.SelectedItems[0];
                    if (connection.State != ConnectionState.Open) connection.Open();
                    command.Connection = connection;
                    command.CommandText = "update " + zvolenaDB + " set Datum=@datum,Castka=@castka,Poznamka=@poznamka" + " where ID=" + radek["ID"];

                    command.Parameters.AddWithValue("@datum", datumVstup);
                    command.Parameters.AddWithValue("@castka", castkaVstup);
                    command.Parameters.AddWithValue("@poznamka", poznamkaVstup);

                    command.ExecuteNonQuery();
                    NaplnTabulkuDatyZeZvoleneDT(zvolenaDB, zvolenaDBZobr);

                    DGrDatabaze.SelectedItems.Clear();
                    BTPridatZaznam.Visibility = Visibility.Visible;
                    BTUpravitZaznam.Visibility = Visibility.Hidden;

                    AktualizujStavCelkovehoStavuFinanci();

                    if (connection.State != ConnectionState.Closed) connection.Close();
                }
                else MessageBox.Show("Zadané vstupní hodnoty neprošly validací. Zkuste hodnoty zadat znovu.", "Chyba", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BTVymazZaznam_Click(object sender, RoutedEventArgs e)
        {
            if (DGrDatabaze.SelectedItems.Count > 0)
            {
                DataRowView radek = (DataRowView)DGrDatabaze.SelectedItems[0];
                if (connection.State != ConnectionState.Open) connection.Open();
                command.Connection = connection;
                command.CommandText = "delete from " + zvolenaDB + " where ID=" + radek["ID"].ToString();
                command.ExecuteNonQuery();

                NaplnTabulkuDatyZeZvoleneDT(zvolenaDB, zvolenaDBZobr);
                DGrDatabaze.SelectedItems.Clear();
                BTPridatZaznam.Visibility = Visibility.Visible;
                BTUpravitZaznam.Visibility = Visibility.Hidden;
                AktualizujStavCelkovehoStavuFinanci();

                if (connection.State != ConnectionState.Closed) connection.Close();
            }
            else MessageBox.Show("Není vybrán žádný záznam ke smazání.", "Varování", MessageBoxButton.OK, MessageBoxImage.Warning);

        }

        private void BTZrusit_Click(object sender, RoutedEventArgs e)
        {
            if (DGrDatabaze.SelectedItems.Count > 0)
            {
                DGrDatabaze.SelectedItems.Clear();

                KrytiMoznostiPriVyberuPrazdnehoRadku(false);
                BTUpravitZaznam.Visibility = Visibility.Hidden;
            }
            VymazDataZFormu();
        }

        private void DGrDatabaze_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BTPridatZaznam.Visibility = Visibility.Visible;
            BTUpravitZaznam.Visibility = Visibility.Visible;

            if (DGrDatabaze.SelectedItems.Count > 0)
            {
                if (DGrDatabaze.SelectedItems[0] is DataRowView radek)
                {
                    KrytiMoznostiPriVyberuPrazdnehoRadku(false);
                    DPDatum.SelectedDate = (DateTime)radek["Datum"];
                    TxBxCastka.Text = radek["Částka"].ToString();

                    TxBxPoznamka.Text = radek["Poznámka"].ToString();
                }
                else
                {
                    KrytiMoznostiPriVyberuPrazdnehoRadku(true);
                }
            }
        }

        private void KrytiMoznostiPriVyberuPrazdnehoRadku(bool kryti)
        {
            if (kryti)
            {
                BTPridatZaznam.Visibility = Visibility.Hidden;
                BTUpravitZaznam.Visibility = Visibility.Hidden;
                BTVymazZaznam.Visibility = Visibility.Hidden;
            }
            else
            {
                BTPridatZaznam.Visibility = Visibility.Visible;
                BTUpravitZaznam.Visibility = Visibility.Visible;
                BTVymazZaznam.Visibility = Visibility.Visible;
            }
        }

        private void BTPrihlasit_Click(object sender, RoutedEventArgs e)
        {
            ProcesPrihlaseni();
        }

        private void ProcesPrihlaseni()
        {
            if (PrihlaseniKDB(PwBxHeslo.Password))
            {
                if (NacistDefaultniNastaveni())
                {
                    PoLoginu();

                    GRLogin.Visibility = Visibility.Hidden;
                    GRProgram.Visibility = Visibility.Visible;
                }
            }
        }

        private void BTUkoncitAplikaci_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void PwBxHeslo_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ProcesPrihlaseni();
            }
        }

        private bool PrihlaseniKDB(string pw)
        {
            try
            {
                if (NacistDefaultniNastaveni())
                {
                    connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={cestaDB}; Jet OLEDB:Database Password={pw};");
                    if (connection.State != ConnectionState.Open) connection.Open();
                    //if (connection.State != ConnectionState.Closed) connection.Close();
                    return true;
                }
                return false;
            }
            catch (OleDbException e)
            {
                if (e.ErrorCode == -2147467259)
                {
                    MessageBox.Show("Nepodařilo se nalézt databázi. Zkontrolujte, zda se databáze nachází ve složce programu, a akci opakujte.", "Chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (e.ErrorCode == -2147217843)
                {
                    MessageBox.Show($"Nepodařilo se přihlásit k databázi. Zkontrolujte přihlašovací údaje a akci opakujte.", "Chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show($"Vyskytla se neočekávaná chyba spojena s připojením k databázi, kvůli které se program ukončí (chybový kód: {e.ErrorCode} - {e.Message})", "Fatální chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                    Close();
                }

                return false;
            }
            catch (Exception e)
            {
                string innerEx = e.InnerException == null ? "-" : e.InnerException.ToString();
                MessageBox.Show($"Vyskytla se neočekávaná chyba, kvůli které se program ukončí." +
                    $"\n\rInner Exception: {innerEx} " +
                    $"\n\rException Message: {e.Message}", "Fatální chyba", MessageBoxButton.OK, MessageBoxImage.Error);
                Close();
                return false;
            }
            finally
            {
                if (connection != null) if (connection.State != ConnectionState.Closed) connection.Close();
                PwBxHeslo.Password = String.Empty;
            }
        }

        private void OdhlaseniUzivatele()
        {
            if (connection.State != ConnectionState.Closed) connection.Close();
            connection = null;
            command = null;
            zvolenaDB = null;
            zvolenaDBZobr = null;
            defaultDB = null;
            defaultZobrNazevDB = null;
            datumVstup = null;
            castkaVstup = 0;
            poznamkaVstup = null;

            GRLogin.Visibility = Visibility.Visible;
            GRProgram.Visibility = Visibility.Hidden;
        }
    }
}
