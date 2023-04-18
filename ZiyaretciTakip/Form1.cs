using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;


namespace ZiyaretciTakip
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Text = DateTime.Now.ToShortDateString();
            dateTimePicker2.Text = DateTime.Now.ToShortDateString();
            if (!File.Exists(@"C:\ZTDB\ZTDB.sqlite"))
            {
                if (!Directory.Exists("C:\\ZTDB\\"))
                {
                    Directory.CreateDirectory("C:\\ZTDB\\");
                    
                }
                DBCreator db = new DBCreator();
                db.veriSetiOlustur();
            }
            RaporVeriGetir();
            ZiyaretciGetir();
            FirmaGetir();
            KartNoGetir();
        }

        public void RaporVeriGetir()
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            con.Open();
            SQLiteDataAdapter da = new SQLiteDataAdapter("select AdSoyad as 'AD SOYAD',TCno as 'TC Kimlik Numarasý',girisTarih as 'Giriþ Tarihi',cikisTarih as 'Çýkýþ Tarihi',FirmaAd as 'Firma Ýsmi',KartNo as 'Kart Numarasý' from Girisler  left join Firmalar f on f.ID=Girisler.firmaID left join Kartlar k on k.ID=Girisler.kartID where Girisler.silindi <>1", con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgwRaporlar.DataSource = dt;
            con.Close();
        }
        public void ZiyaretciGetir()
        {
            GirislerID.Clear();
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            con.Open();
            SQLiteDataAdapter da = new SQLiteDataAdapter("select AdSoyad as 'AD SOYAD',TCno as 'TC Kimlik Numarasý',girisTarih as 'Giriþ Tarihi',cikisTarih as 'Çýkýþ Tarihi',FirmaAd as 'Firma Ýsmi',KartNo as 'Kart Numarasý',Girisler.ID from Girisler  left join Firmalar f on f.ID=Girisler.firmaID left join Kartlar k on k.ID=Girisler.kartID where Girisler.silindi <>1 order by Girisler.ID DESC LIMIT 100", con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgwGiris.DataSource = dt;
            con.Close();
            dgwGiris.Columns[6].Visible = false;
            if (dgwGiris.Rows.Count>0)
            {
                for (int i = 0; i < dgwGiris.Rows.Count; i++)
                {
                    GirislerID.Add(Convert.ToInt32(dgwGiris.Rows[i].Cells[6].Value));
                }
            }
        }

        private List<int> GirislerID = new List<int>();
        private List<int> FirmalarID = new List<int>();
        private List<int> KartlarID = new List<int>();
        public void FirmaGetir()
        {
            cmbFirma.Items.Clear();
            cmbRaporlarFirma.Items.Clear();
            FirmalarID.Clear();
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            con.Open();
            SQLiteDataAdapter da1 = new SQLiteDataAdapter("select FirmaAd as 'Firma Adý',ID from Firmalar", con);
            DataTable dt1 = new DataTable();
            da1.Fill(dt1);
            dgwFirmalar.DataSource = dt1;
            if (dgwFirmalar.Rows.Count>0)
            {
                cmbRaporlarFirma.Items.Add("TÜMÜ");
                for (int i = 0; i < dgwFirmalar.Rows.Count; i++)
                {
                    cmbRaporlarFirma.Items.Add(dgwFirmalar.Rows[i].Cells[0].Value.ToString());
                }
                cmbRaporlarFirma.SelectedIndex = 0;
                cmbRaporlarFirma.Enabled = true;
            }
            else
            {
                cmbRaporlarFirma.Enabled = false; label7.Visible = true;
            }
            SQLiteDataAdapter da = new SQLiteDataAdapter("select FirmaAd as 'Firma Adý',ID from Firmalar where Firmalar.silindi <>1", con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgwFirmalar.DataSource = dt;
            con.Close();
            dgwFirmalar.Columns[1].Visible = false;
            if (dgwFirmalar.Rows.Count>0)
            {
                for (int i = 0; i < dgwFirmalar.Rows.Count; i++)
                {
                    cmbFirma.Items.Add(dgwFirmalar.Rows[i].Cells[0].Value.ToString());
                    FirmalarID.Add(Convert.ToInt32(dgwFirmalar.Rows[i].Cells[1].Value));
                }
                cmbFirma.SelectedIndex = 0;
               
                cmbFirma.Enabled = true;
                label5.Visible = false;
                label7.Visible = false;
            }
            else
            {
                cmbFirma.Enabled = false;
                label5.Visible = true;
            }
        }

        public void KartNoGetir()
        {
            cmbKart.Items.Clear();
            KartlarID.Clear();
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            con.Open();
            SQLiteDataAdapter da = new SQLiteDataAdapter("select KartNo as 'Kart No',ID from Kartlar where silindi <> 1", con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dgwKartlar.DataSource = dt;
            con.Close();
            if (dgwKartlar.Rows.Count > 0)
            {
                for (int i = 0; i < dgwKartlar.Rows.Count; i++)
                {
                    cmbKart.Items.Add(dgwKartlar.Rows[i].Cells[0].Value.ToString());
                    KartlarID.Add(Convert.ToInt32(dgwKartlar.Rows[i].Cells[1].Value));
                }
                cmbKart.SelectedIndex = 0;
                cmbKart.Enabled = true;
                label6.Visible = false;
            }
            else
            {
                cmbKart.Enabled = false;
                label6.Visible = true;
            }
            dgwKartlar.Columns[1].Visible = false;
        }

        private void txtAdSoyad_TextChanged(object sender, EventArgs e)
        {
            if (txtAdSoyad.Text!="")
            {
                panel8.BackColor = Color.Green;
            }
            else
            {
                panel8.BackColor=Color.Red;
            }
            if (panel8.BackColor == Color.Green & panel9.BackColor == Color.Green&cmbFirma.Items.Count>0&cmbKart.Items.Count>0)
            {
                btnZiyaretciEkle.Enabled = true;
            }
            else
            {
                btnZiyaretciEkle.Enabled = false;
            }
        }

        private void txtKimlikNo_TextChanged(object sender, EventArgs e)
        {
            if (txtKimlikNo.Text != "")
            {
                panel9.BackColor = Color.Green;
            }
            else
            {
                panel9.BackColor = Color.Red;
            }

            if (panel8.BackColor == Color.Green & panel9.BackColor == Color.Green & cmbFirma.Items.Count > 0 & cmbKart.Items.Count > 0)
            {
                btnZiyaretciEkle.Enabled = true;
            }
            else
            {
                btnZiyaretciEkle.Enabled = false;
            }
        }

        private void txtFirmaEkle_TextChanged(object sender, EventArgs e)
        {
            if (txtFirmaEkle.Text!="")
            {
                panel10.BackColor=Color.Green;
                btnFirmaEkle.Enabled=true;
            }
            else
            {
                panel10.BackColor=Color.Red;
                btnFirmaEkle.Enabled = false;
            }
        }

        private void txtKartEkle_TextChanged(object sender, EventArgs e)
        {
            if (txtKartEkle.Text!="")
            {
                panel11.BackColor=Color.Green;
                btnKartEkle.Enabled=true;
            }
            else
            {
                panel11.BackColor = Color.Red;
                btnKartEkle.Enabled=false;
            }
        }

        private void btnZiyaretciEkle_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("INSERT INTO Girisler (AdSoyad,TCno,girisTarih,firmaID,kartID,silindi) VALUES ($adsoyad,$tcno,$gTarih,$firma,$kart,0)", con);
            cmd.Parameters.AddWithValue("$adsoyad", txtAdSoyad.Text);
            cmd.Parameters.AddWithValue("$tcno", txtKimlikNo.Text);
            cmd.Parameters.AddWithValue("$gTarih", DateTime.Now.ToString());
            cmd.Parameters.AddWithValue("$firma", FirmalarID[cmbFirma.SelectedIndex]);
            cmd.Parameters.AddWithValue("$kart", KartlarID[cmbKart.SelectedIndex]);
            con.Open();
            try
            {
                    cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            ZiyaretciGetir();
            RaporVeriGetir();
            txtAdSoyad.Text = "";
            txtKimlikNo.Text = "";
            tabControl1.Focus();
        }

        private int GirislerSeciliSatir;

      
        private void çýkýþYapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("update Girisler set cikisTarih=$cTarih where id=$id", con);
            cmd.Parameters.AddWithValue("$cTarih", DateTime.Now.ToString());
            cmd.Parameters.AddWithValue("$id", GirislerID[GirislerSeciliSatir]);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            ZiyaretciGetir();
        }

        private void silToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("update Girisler set silindi=1 where id=$id", con);
            cmd.Parameters.AddWithValue("$id", GirislerID[GirislerSeciliSatir]);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            ZiyaretciGetir();
        }


        private int FirmalarSeciliSatir;

        private void silToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("update Firmalar set silindi=1 where id=$id", con);
            cmd.Parameters.AddWithValue("$id", FirmalarID[FirmalarSeciliSatir]);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            FirmaGetir();
        }

        private int KartlarSeciliSatir;
        private void silToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("update Kartlar set silindi=1 where id=$id", con);
            cmd.Parameters.AddWithValue("$id", KartlarID[KartlarSeciliSatir]);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            KartNoGetir();
        }

        private void btnFirmaEkle_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("INSERT INTO Firmalar (FirmaAd,silindi) VALUES ($ad,0)", con);
            cmd.Parameters.AddWithValue("$ad", txtFirmaEkle.Text);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            FirmaGetir();
            txtFirmaEkle.Text = "";
            tabControl1.Focus();
        }

        private void btnKartEkle_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3");
            SQLiteCommand cmd = new SQLiteCommand("INSERT INTO Kartlar (KartNo,silindi) VALUES ($KartNo,0)", con);
            cmd.Parameters.AddWithValue("$KartNo", txtKartEkle.Text);
            con.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            con.Close();
            KartNoGetir();
            txtKartEkle.Text = "";
            tabControl1.Focus();
        }

        private void dgwGiris_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
          contextMenuStripGirislerCikis.Show(Cursor.Position);
          GirislerSeciliSatir = e.RowIndex;
        }

        private void dgwFirmalar_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           contextMenuStripFirmalarSil.Show(Cursor.Position);
           FirmalarSeciliSatir = e.RowIndex;
        }

        private void dgwKartlar_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
            contextMenuStripKartlarSil.Show(Cursor.Position);
            KartlarSeciliSatir = e.RowIndex;
            
        }

        private void dgwRaporlar_DataSourceChanged(object sender, EventArgs e)
        {
            if (dgwRaporlar.Rows.Count>0)
            {
                btnRaporla.Enabled = true;
            }
            else
            {
                btnRaporla.Enabled = false;
            }
        }

       

        private void txtRaporlarAdSoyad_TextChanged(object sender, EventArgs e)
        {

            if (dgwRaporlar.Rows.Count > 0)
            {
                for (int i = 0; i < dgwRaporlar.Rows.Count; i++)
                {
                    DateTime tarih;
                    if (dgwRaporlar.Rows[i].Cells[3].Value != "")
                    {
                        try
                        {
                            tarih = Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[3].Value.ToString()).Date;
                        }
                        catch (Exception exception)
                        {
                            tarih = DateTime.Now;
                        }

                    }
                    else
                    {
                        tarih = DateTime.Now;
                    }
                    if (cmbRaporlarFirma.SelectedIndex == 0)
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                    else
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & cmbRaporlarFirma.Items[cmbRaporlarFirma.SelectedIndex].ToString() == dgwRaporlar.Rows[i].Cells[4].Value.ToString() & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                }
            }
        }

        private void cmbRaporlarFirma_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dgwRaporlar.Rows.Count > 0)
            {
                for (int i = 0; i < dgwRaporlar.Rows.Count; i++)
                {
                    DateTime tarih;
                    if (dgwRaporlar.Rows[i].Cells[3].Value!="")
                    {
                        try
                        {
                            tarih = Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[3].Value.ToString()).Date;
                        }
                        catch (Exception exception)
                        {
                          tarih=DateTime.Now;
                        }
                       
                    }
                    else
                    {
                        tarih=DateTime.Now;
                    }
                    if (cmbRaporlarFirma.SelectedIndex == 0)
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                    else
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & cmbRaporlarFirma.Items[cmbRaporlarFirma.SelectedIndex].ToString() == dgwRaporlar.Rows[i].Cells[4].Value.ToString() & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (Convert.ToDateTime(dateTimePicker1.Value).Date >= Convert.ToDateTime(dateTimePicker2.Value).Date)
            {
                dateTimePicker2.Value = dateTimePicker1.Value.AddDays(1);
            }
            if (dgwRaporlar.Rows.Count > 0)
            {
                for (int i = 0; i < dgwRaporlar.Rows.Count; i++)
                {
                    DateTime tarih;
                    if (dgwRaporlar.Rows[i].Cells[3].Value != "")
                    {
                        try
                        {
                            tarih = Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[3].Value.ToString()).Date;
                        }
                        catch (Exception exception)
                        {
                            tarih = DateTime.Now;
                        }

                    }
                    else
                    {
                        tarih = DateTime.Now;
                    }
                    if (cmbRaporlarFirma.SelectedIndex == 0)
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                    else
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & cmbRaporlarFirma.Items[cmbRaporlarFirma.SelectedIndex].ToString() == dgwRaporlar.Rows[i].Cells[4].Value.ToString() & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                }
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (Convert.ToDateTime(dateTimePicker2.Value).Date<=Convert.ToDateTime(dateTimePicker1.Value).Date)
            {
                dateTimePicker1.Value = dateTimePicker2.Value.AddDays(-1);
            }
            if (dgwRaporlar.Rows.Count > 0)
            {
                for (int i = 0; i < dgwRaporlar.Rows.Count; i++)
                {
                    DateTime tarih;
                    if (dgwRaporlar.Rows[i].Cells[3].Value != "")
                    {
                        try
                        {
                            tarih = Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[3].Value.ToString()).Date;
                        }
                        catch (Exception exception)
                        {
                            tarih = DateTime.Now;
                        }

                    }
                    else
                    {
                        tarih = DateTime.Now;
                    }
                    if (cmbRaporlarFirma.SelectedIndex == 0)
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                    else
                    {
                        if (dgwRaporlar.Rows[i].Cells[0].Value.ToString().StartsWith(txtRaporlarAdSoyad.Text) & cmbRaporlarFirma.Items[cmbRaporlarFirma.SelectedIndex].ToString() == dgwRaporlar.Rows[i].Cells[4].Value.ToString() & Convert.ToDateTime(dateTimePicker1.Value).Date <= Convert.ToDateTime(dgwRaporlar.Rows[i].Cells[2].Value.ToString()).Date & Convert.ToDateTime(dateTimePicker2.Value).Date >= tarih.Date)
                        {

                            dgwRaporlar.Rows[i].Visible = true;
                        }
                        else
                        {

                            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgwRaporlar.DataSource];
                            currencyManager1.SuspendBinding();
                            dgwRaporlar.Rows[i].Visible = false;
                            currencyManager1.ResumeBinding();
                        }
                    }
                }
            }
        }

        private void btnRaporla_Click(object sender, EventArgs e)
        {
            if (dgwRaporlar.Rows.Count > 0)
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn sutun in dgwRaporlar.Columns)
                {
                    if (sutun.Visible==true)
                    {
                        dt.Columns.Add(sutun.HeaderText);
                    }
                }

                foreach (DataGridViewRow satir in dgwRaporlar.Rows)
                {
                    if (satir.Visible==true)
                    {
                        dt.Rows.Add();
                        foreach (DataGridViewCell hucre in satir.Cells)
                        {
                            dt.Rows[dt.Rows.Count - 1][hucre.ColumnIndex] = hucre.Value.ToString();
                        }
                    }
                }

                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                      "\\Ziyaretçi Takip Raporlar\\"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                              "\\Ziyaretçi Takip Raporlar\\");
                }
                string tarih = DateTime.Now.ToString();
                tarih = tarih.Replace(".", "-");
                tarih = tarih.Replace(":", "-");
                string dosyayolu = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                                   "\\Ziyaretçi Takip Raporlar\\" + tarih + " TARÝHLÝ RAPOR" + ".xlsx";
                using (XLWorkbook wb = new XLWorkbook())
                {
                    
                  
                    wb.Worksheets.Add(dt, "RAPOR");
                    wb.SaveAs(dosyayolu);
                }

                try
                {
                    using (Process myProcess = new Process())
                    {
                        myProcess.StartInfo.UseShellExecute = true;
                        // You can start any process, HelloWorld is a do-nothing example.
                        myProcess.StartInfo.FileName = dosyayolu;
                        myProcess.StartInfo.CreateNoWindow = true;
                        myProcess.Start();
                        // This code assumes the process you are starting will terminate itself.
                        // Given that it is started without a window so you cannot terminate it
                        // on the desktop, it must terminate itself or you can do it programmatically
                        // from this application using the Kill method.
                    }

                    tabControl1.Focus();
                }
                catch (Exception exception)
                {
                   
                }
            }
            
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex==3)
            {
                RaporVeriGetir();
            }
        }
    }
}