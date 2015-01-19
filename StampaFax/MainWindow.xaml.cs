using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Printing;
using System.Windows.Markup;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Xml;
using StampaFax.Presenter;

namespace StampaFax
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        XmlConfig xmlConfig = new XmlConfig();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //per evitare i casini con i path
            try
            {
                string dir = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
                Directory.SetCurrentDirectory(dir);
                Carica();
            }
            catch (Exception ex)
            {
                MessageController.Errore(ex.Message);
                this.Close();
            }
        }

        private void Carica()
        {
            //INIZIALIZZO E CARICO GLI ORDINI DAL FILE EXCEL
            Ordine ordine = new Ordine();

            try
            {
                int numeroRichiesta = 0;                //NUMERO DELLA RICHIESTA
                DateTime dataOrdine = new DateTime();   //DATA DELLA PRELEVATA
                string appValue = "";                   //VALORE DA PASSARE ALLA INPUTBOX

                //CARICO IL NUMERO DELLA RICHIESTA
                do
                {
                    System.Windows.Forms.DialogResult dr = InputBox.ShowInputBox("Inserire il numero ordine: \n(ultimo ordine inserito)", "Numero ordine", ref appValue, xmlConfig.UltimoOrdine.ToString());

                    if (dr == System.Windows.Forms.DialogResult.Cancel) return;

                } while (int.TryParse(appValue, out numeroRichiesta) == false);

                //CARICO LA DATA DELL'ORDINE
                do
                {
                    System.Windows.Forms.DialogResult dr = InputBox.ShowInputBox("Data prelievo: \n(giorno probabile)", "Giorno prelievo", ref appValue, xmlConfig.GetProssimoDataPrelievo());

                    if (dr == System.Windows.Forms.DialogResult.Cancel) return;

                } while (DateTime.TryParse(appValue, out dataOrdine) == false);

                //COPIO E SALVO LA DATA DELLA RICHIESTA
                xmlConfig.UltimoOrdine = numeroRichiesta;

                //STAMPO LE PAGINE
                FixedDocument fix = InizializzaPagina();
                bool finito = StampaPagina(fix, ordine, 1, numeroRichiesta, dataOrdine);
                if (!finito) StampaPagina(fix, ordine, 2, numeroRichiesta, dataOrdine);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private FixedDocument InizializzaPagina()
        {
            try
            {
                FixedDocument fix = new FixedDocument();
                document.Document = fix;
                return fix;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);   
            }

            return null;
        }

        private bool StampaPagina(FixedDocument fixDocument , Ordine ordine , int pagina , int numeroRichiesta , DateTime dataOrdine)
        {
            try
            {
                bool finito = false;

                FixedPage fixedPage = new FixedPage();

                StampaIntestazione(fixDocument, fixedPage, pagina, numeroRichiesta , dataOrdine);
                finito = StampaContenuto(ordine, fixDocument, fixedPage, pagina);
                StampaTotale(fixDocument, fixedPage, pagina, ordine);

                return finito;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return false;
        }

        private void StampaIntestazione(FixedDocument fix, FixedPage page, int pagina , int numeroRichiesta , DateTime dataOrdine)
        {
            TextBlock TB = new TextBlock();

            page.Width = fix.DocumentPaginator.PageSize.Width;
            page.Height = fix.DocumentPaginator.PageSize.Height;

            ImageBrush ib =  new ImageBrush(new BitmapImage( new Uri(
                "pack://application:,,,/StampaFax;component/Images/u88.jpg", UriKind.Absolute)));
            //page.Background = new ImageBrush(new BitmapImage(new Uri("u88.jpg", UriKind.Relative)));
            page.Background = ib;

            //CONVERTO IL NUMERO DELLA RICHIESTA
            string appConv = numeroRichiesta.ToString();
            string numConv = "";
            appConv = appConv.PadLeft(4, ' ');
            foreach (Char c in appConv)
            {
                numConv += c.ToString();
                numConv += " ";
            }

            // NUMERO ORDINE
            TB = new TextBlock();
            TB.Text = numConv;
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(345, 90, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            // COGNOME
            TB = new TextBlock();
            TB.Text = "PISANI";
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(170, 110, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            // NOME
            TB = new TextBlock();
            TB.Text = "ANNA";
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(170, 135, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            // RIV
            TB = new TextBlock();
            TB.Text = "19       ANDRIA";
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(170, 158, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            // TIPOLOGIOA ORDINE
            TB = new TextBlock();
            TB.Text = "";
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(555, 130, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            // PAGINA
            TB = new TextBlock();
            TB.Text = pagina + "  2";
            TB.FontSize = 23; // 30pt text
            TB.Margin = new Thickness(695, 130, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);

            //CONVERTO LA DATA
            string giorno = ((int)dataOrdine.Day).ToString("00");
            giorno = giorno.Insert(1, " ");
            string mese = ((int)dataOrdine.Month).ToString("00");
            mese = mese.Insert(1, " ");
            string anno = ((int)dataOrdine.Year).ToString().Substring ( 2 , 2 );
            anno = anno.Insert(1, " ");


            // DATA CONSEGNA - ELIMINATO
            TB = new TextBlock();
            TB.Text = giorno + "  " + mese + "  " + anno;
            TB.FontSize = 22; // 30pt text
            TB.Margin = new Thickness(555, 158, 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            //page.Children.Add(TB);


            // add the page to the document
            PageContent pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(page);
            fix.Pages.Add(pageContent);
        }

        private bool StampaContenuto(Ordine ordine , FixedDocument fix, FixedPage page , int pagina)
        {
            const int _RECORDPERFOGLIO = 48;
            TextBlock TB = new TextBlock();
            double riga = 0;
            int i = (pagina - 1) * _RECORDPERFOGLIO;

            //PRIMA COLONNA
            do
            {
                Record r = ordine.getOrdine(i);

                // CODICE
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaCodice(); 
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(110, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                // CHILI
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaChili();
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(222, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                // GRAMMI
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaGrammi();
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(315, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                riga += 27.5;

                i++;

            } while (i < ordine.getNumeroRecord() && i < _RECORDPERFOGLIO/2 + (pagina - 1) * _RECORDPERFOGLIO);

            if (i == ordine.getNumeroRecord()) return true;

            //SECONDA COLONNA

            riga = 0;
            double colonna = 346;

            do
            {
                Record r = ordine.getOrdine(i);

                // CODICE
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaCodice(); 
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(110 + colonna, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                // CHILI
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaChili();
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(222 + colonna, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                // GRAMMI
                TB = new TextBlock();
                TB.ToolTip = r.nome;
                TB.Text = r.StampaGrammi();
                TB.FontSize = 22; // 30pt text
                TB.Margin = new Thickness(315 + colonna, 233 + riga, 100, 0); // 1 inch margin
                TB.FontFamily = new FontFamily("Consolas");
                page.Children.Add(TB);

                riga += 27.5;

                i++;

            } while (i < ordine.getNumeroRecord() && i < _RECORDPERFOGLIO + (pagina - 1) * _RECORDPERFOGLIO);

            if (i == ordine.getNumeroRecord()) return true;

            return false;
        }        

        private void StampaTotale(FixedDocument fix, FixedPage page, int pagina , Ordine ordine)
        {
            TextBlock TB = new TextBlock();

            // GRAMMI
            TB = new TextBlock();
            TB.Text = ordine.getChiliTot(pagina) + "  " + ordine.getGrammiTot(pagina);
            TB.FontSize = 22; // 30pt text
            TB.Margin = new Thickness(466, 903 , 100, 0); // 1 inch margin
            TB.FontFamily = new FontFamily("Consolas");
            page.Children.Add(TB);
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            FrmImpostaGiorno frm = new FrmImpostaGiorno(xmlConfig);
            frm.Show();
        }

        private void cmdRicarica_Click(object sender, RoutedEventArgs e)
        {
            Carica();
        }

    }
}
