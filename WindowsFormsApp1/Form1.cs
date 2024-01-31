using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Word.Application winword = new Word.Application();
        Word.Document document;
        double ResPrice = 0;
        List<double> ShopLinePrice = new List<double>();
        List<Commodity> Shop = new List<Commodity>()
        {
            new Commodity(){ Id = "1", Product = "Апельсины", Count = "50", Price = "120"},
            new Commodity(){ Id = "2", Product = "Арбузы", Count = "95", Price = "230"},
            new Commodity(){ Id = "3", Product = "Тыква", Count = "140", Price = "150"},
            new Commodity(){ Id = "4", Product = "Кокосы", Count = "15", Price = "210"},
            new Commodity(){ Id = "5", Product = "Яблоки", Count = "200", Price = "15"},
            new Commodity(){ Id = "6", Product = "Киви", Count = "100", Price = "125"}
        };
        public Form1()
        {
            InitializeComponent();
            document = winword.Documents.Add();
            label4.Text = "Дата: " + DateTime.Now.ToString("dd.MM.yyyy");
            foreach (Commodity com in Shop)
            {
                double summ = Convert.ToInt32(com.Count) * Convert.ToInt32(com.Price);
                ShopLinePrice.Add(summ);
                ResPrice += summ;
            }
            label5.Text = "Итого: " + ResPrice;
            dataGridView1.ColumnCount = 5;
            dataGridView1.RowCount = Shop.Count;
            dataGridView1.Columns[0].HeaderText = "ID";
            dataGridView1.Columns[1].HeaderText = "Продукт";
            dataGridView1.Columns[2].HeaderText = "Цена";
            dataGridView1.Columns[3].HeaderText = "Количество";
            dataGridView1.Columns[4].HeaderText = "Сумма";
            for (int i = 0; i < Shop.Count; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = Shop[i].Id;
                dataGridView1.Rows[i].Cells[1].Value = Shop[i].Product;
                dataGridView1.Rows[i].Cells[2].Value = Shop[i].Price;
                dataGridView1.Rows[i].Cells[3].Value = Shop[i].Count;
                dataGridView1.Rows[i].Cells[4].Value = ShopLinePrice[i];
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Word.Paragraph invoicePar = document.Content.Paragraphs.Add();
            DateTime? selectDate = DateTime.Now;
            string invoiceNumber = textBox3.Text;
            if (selectDate != null)
            {
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber, " от ", selectDate.Value.ToString("dd.MM.yyyy"));
            }
            else
            {
                invoicePar.Range.Text = string.Concat("Расходная накладная № ", invoiceNumber);
            }
            invoicePar.Range.Font.Name = "Times new roman";
            invoicePar.Range.Font.Size = 14;
            invoicePar.Range.InsertParagraphAfter();

            string PurchasertxtBox = textBox1.Text;
            Word.Paragraph providerPar = document.Content.Paragraphs.Add();
            providerPar.Range.Text = string.Concat("Поставщик: ", PurchasertxtBox);
            providerPar.Range.Font.Name = "Times new roman";
            providerPar.Range.Font.Size = 14;
            providerPar.Range.InsertParagraphAfter();

            Word.Paragraph customerPar = document.Content.Paragraphs.Add();
            string ProvidertxtBox = textBox2.Text;
            customerPar.Range.Text = "Покупатель: " + ProvidertxtBox;
            customerPar.Range.Font.Name = "Times new roman";
            customerPar.Range.Font.Size = 14;
            customerPar.Range.InsertParagraphAfter();


            int nRows = Shop.Count;
            Word.Table myTable = document.Tables.Add(customerPar.Range, nRows, 5);
            myTable.Borders.Enable = 1;
            for (int i = 1; i < Shop.Count + 1; i++)
            {
                var dataRow = myTable.Rows[i].Cells;
                dataRow[1].Range.Text = Shop[i - 1].Id.ToString();
                dataRow[2].Range.Text = Shop[i - 1].Product;
                dataRow[3].Range.Text = Shop[i - 1].Count.ToString();
                dataRow[4].Range.Text = Shop[i - 1].Price.ToString();
                dataRow[5].Range.Text = ShopLinePrice[i - 1].ToString();
            }
            Word.Paragraph resPar = document.Content.Paragraphs.Add();
            resPar.Range.Text = string.Concat("Итого:", ResPrice);
            resPar.Range.Font.Name = "Times new roman";
            resPar.Range.Font.Size = 14;
            resPar.Range.InsertParagraphAfter();

            object filename = @"H:\wordExample.doc";
            document.SaveAs(filename);
            document.Close();
            winword.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
    public class Commodity
    {
        public string Id;
        public string Product;
        public string Count;
        public string Price;
    }

}