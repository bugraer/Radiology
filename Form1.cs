using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;
namespace derle
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }
        double l=0;
        double sutun1dogru, sutun1say;
        string[] sutun1 = new string[10];
        double[] sutun1d = new double[10];
        double o;


      //  double l1 = 0;
        double sutun2dogru, sutun2say;
       /* string[] sutun2 = new string[10];
        double[] sutun2d = new double[10];
        */
        private void button1_Click(object sender, EventArgs e)
        {
          
            string DosyaYolu;
            string DosyaAdi;
            DataTable dt;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;// seçilen dosyanın tüm yolunu verir
                DosyaAdi = file.SafeFileName;// seçilen dosyanın adını verir.
                ExcelApp.Application excelApp = new ExcelApp.Application();
                if (excelApp == null)
                { //Excel Yüklümü Kontrolü Yapılmaktadır.
                    MessageBox.Show("Excel yüklü değil.");
                    return;
                }

                //Excel Dosyası Açılıyor.
                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(DosyaYolu);
                //Excel Dosyasının Sayfası Seçilir.
                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                //Excel Dosyasının ne kadar satır ve sütun kaplıyorsa tüm alanları alır.
                ExcelApp.Range excelRange = excelSheet.UsedRange;

                int satirSayisi = excelRange.Rows.Count; //Sayfanın satır sayısını alır.
                int sutunSayisi = excelRange.Columns.Count; //Sayfanın sütun sayısını alır.
                dt = ToDataTable(excelRange, satirSayisi, sutunSayisi);

                dataGridView1.DataSource = dt;
                dataGridView1.Refresh();

                //Okuduktan Sonra Excel Uygulamasını Kapatıyoruz.
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            }
            else
            {
                MessageBox.Show("Dosya Seçilemedi.");
            }
        }

        public DataTable ToDataTable(ExcelApp.Range range, int rows, int cols)
        {
            DataTable table = new DataTable();

            for (int i = 1; i <= rows; i++)
            {
                if (i == 1)
                { // ilk satırı Sutun Adları olarak kullanıldığından bunları Sutün Adları Olarak Kaydediyoruz.
                    for (int j = 1; j <= cols; j++)
                    {
                        //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                            table.Columns.Add(range.Cells[i, j].Value2.ToString());
                        else
                            table.Columns.Add(j.ToString() + ".Sütun"); //Boş olduğunda Kaçınsı Sutünsa Adı veriliyor.
                    }
                    continue;
                }

                //Yukarıda Sütunlar eklendi onun şemasına göre yeni bir satır oluşturuyoruz. 
                //Okunan verileri yan yana sıralamak için
                var yeniSatir = table.NewRow();
                for (int j = 1; j <= cols; j++)
                {
                    //Sütunların içeriği boş mu kontrolü yapılmaktadır.
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        yeniSatir[j - 1] = range.Cells[i, j].Value2.ToString();
                    else
                        yeniSatir[j - 1] = String.Empty; // İçeriği boş hücrede hata vermesini önlemek için
                }
                table.Rows.Add(yeniSatir);

            }


            for (int suttun = 6; suttun <= 15; suttun++)
            {
                l = 0;
                o = 0;
                sutun1say = 0;
                    sutun1dogru = 0;
                sutun2say = 0;
                sutun2dogru = 0;
                string[] k = new string[9];
                for (int g = 2; g <= 5; g++)
                {
                    k[g] = range.Cells[g, suttun].Value2.ToString();
                    l = l + Convert.ToDouble(k[g]);

                }
                o = l / 4;


                for (int gg = 2; gg <= 5; gg++)
                {
                    if (range.Cells[gg, 4].Value2.ToString() == "luminal A")
                    {

                        sutun1[gg] = range.Cells[gg, suttun].Value2.ToString();
                        sutun1dogru++;
                        if (Convert.ToDouble(sutun1[gg]) < o)
                        {
                            sutun1say++;
                        }

                    }
                }
                for (int ggg = 2; ggg <= 5; ggg++)
                {
                    if (range.Cells[ggg, 4].Value2.ToString() == "luminal B")
                    {

                        sutun1[ggg] = range.Cells[ggg, suttun].Value2.ToString();
                        sutun2dogru++;
                        if (Convert.ToDouble(sutun1[ggg]) > o)
                        {
                            sutun2say++;
                        }

                    }
                }
                if ((sutun1dogru == sutun1say) && (sutun2dogru == sutun2say))
                {
                    if (suttun == 6)
                    {
                        label1.Text = range.Cells[1, 6].Value2.ToString();
                    }
                    if (suttun == 7)
                    {
                        label2.Text = range.Cells[1, 7].Value2.ToString();
                    }
                    if (suttun == 8)
                    {
                        label3.Text = range.Cells[1, 8].Value2.ToString();
                    }
                    if (suttun ==9)
                    {
                        label4.Text = range.Cells[1, 9].Value2.ToString();
                    }
                    if (suttun == 10)
                    {
                        label5.Text = range.Cells[1, 10].Value2.ToString();
                    }
                    if (suttun == 11)
                    {
                        label6.Text = range.Cells[1, 11].Value2.ToString();
                    }
                    if (suttun == 12)
                    {
                        label7.Text = range.Cells[1, 12].Value2.ToString();
                    }
                    if (suttun == 13)
                    {
                        label8.Text = range.Cells[1, 13].Value2.ToString();
                    }
                    if (suttun == 14)
                    {
                        label9.Text = range.Cells[1, 14].Value2.ToString();
                    }
                    if (suttun == 15)
                    {
                        label10.Text = range.Cells[1, 15].Value2.ToString();
                    }
                }
                else
                {
                    if (suttun == 6)
                    {
                        label1.Text = "";
                    }
                    if (suttun == 7)
                    {
                        label2.Text = "";
                    }
                    if (suttun == 8)
                    {
                        label3.Text = "";
                    }
                    if (suttun == 9)
                    {
                        label4.Text = "";
                    }
                    if (suttun == 10)
                    {
                        label5.Text = "";
                    }
                    if (suttun == 11)
                    {
                        label6.Text = "";
                    }
                    if (suttun == 12)
                    {
                        label7.Text = "";
                    }
                    if (suttun == 13)
                    {
                        label8.Text = "";
                    }
                    if (suttun == 14)
                    {
                        label9.Text = "";
                    }
                    if (suttun == 15)
                    {
                        label10.Text = "";
                    }
                }
               // label14.Text = String.Format(Math.Round(l, 0).ToString());
            }
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   /*         string[] k1 = new string[9];
            for (int g = 2; g <= 5; g++)
            {
                k1[g] = range.Cells[g, 8].Value2.ToString();
                l1 = l1 + Convert.ToDouble(k1[g]);

            }
            double o1 = l1 / 4;
            label3.Text = String.Format(Math.Round(o1, 0).ToString());

            for (int g = 2; g <= 5; g++)
            {
                if (range.Cells[g, 4].Value2.ToString() == "luminal A")
                {

                    sutun2[g] = range.Cells[g, 8].Value2.ToString();
                    sutun2dogru++;
                    if (Convert.ToDouble(sutun2[g]) < l)
                    {
                        sutun2say++;
                    }

                }
            }
            if (sutun2dogru == sutun2say)
            {
                label13.Text = "bumbe yarag";
            }



*/


            /*



                            string[] x = new string[4];
                        string[] y = new string[4];
                        string[] z = new string[4];
                        int[] AA = new int[4];
                        int[] BB = new int[4];
                        {
                            x[0] = range.Cells[2, 6].Value2.ToString();
                            x[1] = range.Cells[3, 6].Value2.ToString();
                            x[2] = range.Cells[4, 6].Value2.ToString();
                            x[3] = range.Cells[5, 6].Value2.ToString();

                            string abc = range.Cells[3, 4].Value2.ToString();
                            if(abc == "luminal B")
                            {
                                label12.Text = "5";
                            }
                            label13.Text = range.Cells[1, 4].Value2.ToString();
                        }
                        double a = Convert.ToDouble(x[0]) + Convert.ToDouble(x[1]) + Convert.ToDouble(x[2]) + Convert.ToDouble(x[3]);
                        a = a / 4;
            label1.Text = String.Format(Math.Round(a, 0).ToString());

                        for (int kkk = 2; kkk <= 5; kkk++)
                        {

                        }


                        label5.Text = String.Format(k[2].ToString());
                        label7.Text = String.Format(k[3].ToString());
                        label9.Text = String.Format(k[4].ToString());

                        */
            return table;
        }

      

    
        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
    }

