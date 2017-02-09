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
    class ModifierParametres
    {
        //modifier Ministere
        public void modifier_ministrere(String ch)
        {
            // SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {
                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET Ministere='" + ch + "' ", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("veillez doubler les apstrophes que vous avez mise"); }
            con.seDeconnecter();


        }
        //modifier Organisme
        public void modifier_organisme(String ch)
        {


            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {
                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET Organisme='" + ch + "'", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("veillez doubler les apstrophes que vous avez mise"); }
            con.seDeconnecter();

        }
        //modifier compteSoEsi
        public void modifier_CompteSocEsi(String ch)
        {

            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {

                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET CompteSocEsi='" + ch + "' ", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("veillez doubler les apstrophes que vous avez mise"); }

            con.seDeconnecter();
        }












        //Modifier compteEsiTresor
        public void modifier_CompteEsiTresor(String ch)
        {
            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {

                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET CompteEsiTresor='" + ch + "' ", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("veillez doubler les apstrophes que vous avez mise"); }
            con.seDeconnecter();
        }


        //modifier la durée de cotisation 
        public void modifier_cotisation(int ch)
        {
            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {

                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET DurCot=" + ch + " ", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("Attention la sécie doit contenir que des chiffres!!"); }
            con.seDeconnecter();


        }
        //modifier le mois début de l'année sociale
        public void modifier_mois(int ch)
        {
            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {

                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET MoisDebAnSoc=" + ch + " ", con.connextion() );
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("Attention la sécie doit contenir que des chiffres!!"); }
            con.seDeconnecter();


        }

        //modifer le jour début de l'année sociale
        public void modifier_jour(int ch)
        {
            //SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=OeuvresSociales;Persist Security Info=True;User ID=sa;Password=mama0552331423");
            BD con = new BD();
            try
            {

                con.seConnecter();
                SqlCommand cmd = new SqlCommand("UPDATE Parametres SET JourDebAnSoc=" + ch + " ", con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show("modification effectué");
            }
            catch (SqlException ex)
            { MessageBox.Show("Attention la sécie doit contenir que des chiffres!!"); }
            con.seDeconnecter();


        }


        //*************Fin*****************//




    }
}
