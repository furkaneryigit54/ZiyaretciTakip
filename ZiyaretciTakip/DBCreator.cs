using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZiyaretciTakip
{
    internal class DBCreator
    {

        public  void veriSetiOlustur()
        {
            //veritabanındaki tabloları oluşturacak sorgular
            string[] sorgular =
            {
                "CREATE TABLE \"Firmalar\" (\r\n\t\"ID\"\tINTEGER NOT NULL UNIQUE,\r\n\t\"FirmaAd\"\tTEXT,\r\n\t\"silindi\"\tINTEGER,\r\n\tPRIMARY KEY(\"ID\" AUTOINCREMENT)\r\n);",
                "CREATE TABLE \"Girisler\" (\r\n\t\"ID\"\tINTEGER NOT NULL UNIQUE,\r\n\t\"AdSoyad\"\tTEXT,\r\n\t\"TCno\"\tTEXT,\r\n\t\"girisTarih\"\tTEXT,\r\n\t\"cikisTarih\"\tTEXT,\r\n\t\"firmaID\"\tINTEGER,\r\n\t\"kartID\"\tINTEGER,\r\n\t\"silindi\"\tINTEGER,\r\n\tPRIMARY KEY(\"ID\" AUTOINCREMENT)\r\n);",
                "CREATE TABLE \"Kartlar\" (\r\n\t\"ID\"\tINTEGER NOT NULL UNIQUE,\r\n\t\"KartNo\"\tTEXT,\r\n\t\"silindi\"\tINTEGER,\r\n\tPRIMARY KEY(\"ID\" AUTOINCREMENT)\r\n);"
            };
            query(sorgular);
        }
        static void query(string[] sorgular)
        {
            //sorgu metinlerini işleme metotu
            string dbdosya = "Data Source=\"C:\\ZTDB\\ZTDB.sqlite\";Version=3";
            SQLiteConnection con = new SQLiteConnection(dbdosya);
            con.Open();
            foreach (string sorgu in sorgular)
            {
                SQLiteCommand cmd = new SQLiteCommand(sorgu, con);
                cmd.ExecuteNonQuery();
            }
            con.Close();
        }
    }
}
