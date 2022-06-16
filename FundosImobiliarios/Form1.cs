using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using HtmlAgilityPack;
using ClosedXML.Excel;
using System.IO;

namespace FundosImobiliarios
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            mPegarInformacoes();
            this.Close();
        }
        private void mPegarInformacoes() {
            try
            {
                string iHtmlPagina = mPageSouce("https://www.fundsexplorer.com.br/ranking");
                DataTable dtTabela = mHtmlTableToDataTable(iHtmlPagina);
                string iDataAtual = DateTime.Now.ToString("yyyy-MM-dd");
                mDataTableToWorksheet(dtTabela, "C:\\Users\\iago\\Downloads\\FUNDOS IMOBILIARIOS\\", $"FII_{iDataAtual}", "xlsx");
            }
            catch (Exception ex)
            {
                mWriteLog("C:\\Users\\iago\\Downloads\\FUNDOS IMOBILIARIOS\\", "A", $"{DateTime.Now.ToString("yyyy-MM-dd")} - {ex.Message}");
                MessageBox.Show(ex.Message);
            }
        }
        private string mPageSouce(string pURL) {
            using (WebClient iWebClient = new WebClient()) {
                iWebClient.Encoding = Encoding.UTF8;
                return iWebClient.DownloadString(pURL);
            }
        }
        private DataTable mHtmlTableToDataTable(string pHtml)
        {
            DataTable table = new DataTable();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(pHtml);
            var headers = doc.DocumentNode.SelectNodes("//tr/th");
            foreach (HtmlNode header in headers)
                table.Columns.Add(header.InnerText); // create columns from th
                                                     // select rows with td elements 
            foreach (var row in doc.DocumentNode.SelectNodes("//tr[td]"))
                table.Rows.Add(row.SelectNodes("td").Select(td => td.InnerText.Replace(".0",",0").Replace("N/A","0")).ToArray());
            return table;
        }

        private void mDataTableToWorksheet(DataTable pDataTable,string pFilePath,string pWorksheetName, string pFileExtension = "xlsx") {
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(pDataTable, pWorksheetName);
            wb.SaveAs($"{pFilePath}{pWorksheetName}.{pFileExtension}");        
        }

        private void mWriteLog(string pLogPath,string pFileName,string pTextToWrite) {

            if (!File.Exists($"{pLogPath}{pFileName}.txt"))
            {
                File.Create($"{pLogPath}{pFileName}.txt").Close();             
            }
            using (TextWriter writer = new StreamWriter($"{pLogPath}{pFileName}.txt", true))
            {
                writer.WriteLine(pTextToWrite);
            }           
            
        }

    }
    
}
