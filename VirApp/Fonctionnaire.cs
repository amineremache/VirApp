using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;

namespace VirApp
{
    class Fonctionnaire
    {
        private string nom;
        private string pnom;
        private string email;
        private string comptefont;
        private int matricule;
        private long Ntel;
        private string cb;
        public string SitFam;
        DateTime? date;
        public Fonctionnaire(int matricule, string nom, string pnom, DateTime? date, long ntel, string email, string comptfont, string cb, string Sf)
        {
            this.nom = nom;
            this.pnom = pnom;
            this.comptefont = comptfont;
            this.matricule = matricule;
            this.email = email;
            this.date = date;
            this.Ntel = ntel;
            this.cb = cb;
            this.SitFam = Sf;
        }
        public Fonctionnaire()
        {
            this.nom = null;
            this.pnom = null;
            this.comptefont = null;
            //this.matricule = 0;
            this.email = "";
            this.date = new DateTime(0001, 01, 01);
            this.Ntel = 0;
        }

        public void Add_fonctionnaire()
        {

            try
            {
                BD con = new BD();
                con.seConnecter();
                String requete = "INSERT INTO Fonctionnaire (Matricule, NomFonct, PrenFonct, DateRecrut, TelFonct, EmailFonct, CompteFonct, CodeBanque , SitFamFonct) VALUES('" + matricule + "','" + nom + "','" + pnom + "','" + date + "','" + Ntel + "','" + email + "','" + comptefont + "','" + cb + "','" + SitFam + "')";
                SqlCommand cmd = new SqlCommand(requete, con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show(" réussite de l ' addition ");
                con.seDeconnecter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Oups !Ce fonctionnaire existe deja ! \n" + ex.Message);
            }

        }
    }
}
