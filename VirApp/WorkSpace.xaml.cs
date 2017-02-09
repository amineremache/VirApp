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
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.Common;
using System.Data;
using System.IO;
using Microsoft.Win32;


namespace VirApp
{

    public partial class WorkSpace : UserControl
    {
        public WorkSpace()
        {
            InitializeComponent();
            Upload();
            Upload_traitement();
            
            


        }




        // Attributs de la classe



        public static string path = System.IO.Directory.GetCurrentDirectory();
        public static string conString = "Data Source=(LocalDB)\\v11.0; AttachDbFilename=\"" + path + "\\BDD\\OeuvresSociales2.mdf\";Integrated Security=True";

        //public static string conString = "Data Source=localhost;Initial Catalog=OeuvresSociales2;Integrated Security=True";


        public static SqlConnection connect = new SqlConnection(conString);
        public string Query;
        public static SqlCommand command = new SqlCommand();




        static string lastKeyQuery = "SELECT NumDem FROM DemandePrime";
        public static string lastKey = getLastKey(lastKeyQuery);



        // Methodes de la classe 

        public static string GetterLastKey()
        {
            return lastKey;
        }


        public string verifNumDemaAndTypePrime(string NumDemFormulaire)
        {
            Query = "SELECT CodePrime FROM DemandePrime ";
            Query += " WHERE NumDem=" + NumDemFormulaire;

            string CodePrime = getExcuteScalar(Query);

            Query = "SELECT DésignationPrime FROM TypePrime ";
            Query += " WHERE CodePrime=" + CodePrime;

            string DésignationPrime = getExcuteScalar(Query);

            if (DésignationPrime == "Don") return "Don";
            else if (DésignationPrime == "Décès Employée") return "Décès Employée";
            else if (DésignationPrime == "") return "";
            else return "Autres";



        }

        public string getExcuteScalar(string Query)
        {
            SqlCommand cmd = new SqlCommand(Query, connect);
            string result = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }
            //MessageBox.Show(Query);
            if (cmd.ExecuteScalar() != null) result = cmd.ExecuteScalar().ToString();
            else result = null;
            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return result;
        }

        public class ConvertisseurChiffresLettres
        {

            public string convertion(double chiffre)
            {
                int centaine, dizaine, unite, reste, y;
                bool dix = false;
                bool soixanteDix = false;
                string lettre = "";

                reste = (int)chiffre / 1;

                for (int i = 1000000000; i >= 1; i /= 1000)
                {
                    y = reste / i;
                    if (y != 0)
                    {
                        centaine = y / 100;
                        dizaine = (y - centaine * 100) / 10;
                        unite = y - (centaine * 100) - (dizaine * 10);
                        switch (centaine)
                        {
                            case 0:
                                break;
                            case 1:
                                lettre += "cent ";
                                break;
                            case 2:
                                if ((dizaine == 0) && (unite == 0)) lettre += "deux cents ";
                                else lettre += "deux-cent ";
                                break;
                            case 3:
                                if ((dizaine == 0) && (unite == 0)) lettre += "trois cents ";
                                else lettre += "trois-cent ";
                                break;
                            case 4:
                                if ((dizaine == 0) && (unite == 0)) lettre += "quatre cents ";
                                else lettre += "quatre-cent ";
                                break;
                            case 5:
                                if ((dizaine == 0) && (unite == 0)) lettre += "cinq cents ";
                                else lettre += "cinq-cent ";
                                break;
                            case 6:
                                if ((dizaine == 0) && (unite == 0)) lettre += "six cents ";
                                else lettre += "six-cent ";
                                break;
                            case 7:
                                if ((dizaine == 0) && (unite == 0)) lettre += "sept cents ";
                                else lettre += "sept-cent ";
                                break;
                            case 8:
                                if ((dizaine == 0) && (unite == 0)) lettre += "huit cents ";
                                else lettre += "huit-cent ";
                                break;
                            case 9:
                                if ((dizaine == 0) && (unite == 0)) lettre += "neuf cents ";
                                else lettre += "neuf-cent ";
                                break;
                        }// La fin du cas " centaine "

                        switch (dizaine)
                        {
                            case 0:
                                break;
                            case 1:
                                dix = true;
                                break;
                            case 2:
                                lettre += "vingt ";
                                break;
                            case 3:
                                lettre += "trente ";
                                break;
                            case 4:
                                lettre += "quarante ";
                                break;
                            case 5:
                                lettre += "cinquante ";
                                break;
                            case 6:
                                lettre += "soixante ";
                                break;
                            case 7:
                                dix = true;
                                soixanteDix = true;
                                lettre += "soixante ";
                                break;
                            case 8:
                                lettre += "quatre-vingt ";
                                break;
                            case 9:
                                dix = true;
                                lettre += "quatre-vingt ";
                                break;
                        } // La fin du cas " dizaine "

                        switch (unite)
                        {
                            case 0:
                                if (dix) lettre += "dix ";
                                break;
                            case 1:
                                if (soixanteDix) lettre += "et onze ";
                                else
                                    if (dix) lettre += "onze ";
                                else if ((dizaine != 1 && dizaine != 0)) lettre += "et un ";
                                else lettre += "un ";
                                break;
                            case 2:
                                if (dix) lettre += "douze ";
                                else lettre += "quatre ";
                                break;
                            case 5:
                                if (dix) lettre += "quinze ";
                                else lettre += "cinq ";
                                break;
                            case 6:
                                if (dix) lettre += "seize ";
                                else lettre += "six ";
                                break;
                            case 7:
                                if (dix) lettre += "dix-sept ";
                                else lettre += "sept ";
                                break;
                            case 8:
                                if (dix) lettre += "dix-huit ";
                                else lettre += "huit ";
                                break;
                            case 9:
                                if (dix) lettre += "dix-neuf ";
                                else lettre += "neuf ";
                                break;
                        } // La fin du cas " unite "

                        switch (i)
                        {
                            case 1000000000:
                                if (y > 1) lettre += "milliards ";
                                else lettre += "milliard ";
                                break;
                            case 1000000:
                                if (y > 1) lettre += "millions ";
                                else lettre += "million ";
                                break;
                            case 1000:
                                lettre += "mille ";
                                break;
                        }
                    } // la fin de la condition if ( y!= 0 )
                    reste -= y * i;
                    dix = false;
                    soixanteDix = false;
                } // la fin de la boucle "pour" 

                if (lettre.Length == 0) lettre += "zero";

                // pour les chiffres apres la virgule :

                Decimal chiffresDecimals;
                chiffresDecimals = (Decimal)(chiffre * 100) % 100;


                dizaine = (int)(chiffresDecimals) / 10;
                unite = (int)chiffresDecimals - (dizaine * 10);

                string lettreDecimal = "";
                switch (dizaine)
                {
                    case 0:
                        break;
                    case 1:
                        dix = true;
                        break;
                    case 2:
                        lettreDecimal += "vingt ";
                        break;
                    case 3:
                        lettreDecimal += "trente ";
                        break;
                    case 4:
                        lettreDecimal += "quarante ";
                        break;
                    case 5:
                        lettreDecimal += "cinquante ";
                        break;
                    case 6:
                        lettreDecimal += "soixante ";
                        break;
                    case 7:
                        dix = true;
                        soixanteDix = true;
                        lettreDecimal += "soixante ";
                        break;
                    case 8:
                        lettreDecimal += "quatre-vingt ";
                        break;
                    case 9:
                        dix = true;
                        lettreDecimal += "quatre-vingt ";
                        break;
                } // La fin du cas " dizaine "

                switch (unite)
                {
                    case 0:
                        if (dix) lettreDecimal += "dix ";
                        break;
                    case 1:
                        if (soixanteDix) lettreDecimal += "et onze ";
                        else
                            if (dix) lettreDecimal += "onze ";
                        else if ((dizaine != 1 && dizaine != 0)) lettreDecimal += "et un ";
                        else lettreDecimal += "un ";
                        break;
                    case 2:
                        if (dix) lettreDecimal += "douze ";
                        else lettreDecimal += "deux ";
                        break;
                    case 3:
                        if (dix) lettreDecimal += "treize ";
                        else lettreDecimal += "trois ";
                        break;
                    case 4:
                        if (dix) lettreDecimal += "quatorze ";
                        else lettreDecimal += "quatre ";
                        break;
                    case 5:
                        if (dix) lettreDecimal += "quinze ";
                        else lettreDecimal += "cinq ";
                        break;
                    case 6:
                        if (dix) lettreDecimal += "seize ";
                        else lettreDecimal += "six ";
                        break;
                    case 7:
                        if (dix) lettreDecimal += "dix-sept ";
                        else lettreDecimal += "sept ";
                        break;
                    case 8:
                        if (dix) lettreDecimal += "dix-huit ";
                        else lettreDecimal += "huit ";
                        break;
                    case 9:
                        if (dix) lettreDecimal += "dix-neuf ";
                        else lettreDecimal += "neuf ";
                        break;
                } // La fin du cas " unite "


                // Traiter le cas de " un mille " :

                if (lettre.StartsWith("un mille")) lettre = lettre.Remove(0, 3);

                /* Rajouter la devise ( Dinars ) et traitement des cas spéciaux */

                if (lettreDecimal.Equals(""))
                {
                    if (lettre.Equals("un "))
                        return lettre + "dinar";
                    else
                        return lettre + "dinars";
                }
                else if (dizaine.Equals(0) && unite.Equals(1))
                {
                    if (lettre.Equals("un "))
                        return lettre + "dinar et " + lettreDecimal + "centime";
                    else
                        return lettre + "dinars et " + lettreDecimal + "centime";
                }

                else
                    return lettre + "dinars et " + lettreDecimal + "centimes";
            }


            // Methode pour mettre la première lettre en majuscule
            public string PremiereLettreMaj(string ChaineAConvertir)
            {
                if (!(String.IsNullOrEmpty(ChaineAConvertir)))
                {
                    return ChaineAConvertir.First().ToString().ToUpper() + String.Join("", ChaineAConvertir.Skip(1));
                }
                else
                {
                    return ChaineAConvertir;
                }
            }
        }

        static public string getLastKey(string Query)
        {
            string LastKey = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }

            SqlCommand cmd = new SqlCommand(Query, connect);

            SqlDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                LastKey = reader[0].ToString();
            }

            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return LastKey;
        }





        // Partie visuel

        public enum choix
        {
            Mise_a_jour,
            Demande,
            Virement,
            Statistiques
        }


        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (((ListBox)sender).SelectedIndex)
            {
                case 0:
                    accueil.Visibility = Visibility.Visible;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    
                    break;
                case 1:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Visible;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 2:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Visible;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;


                    break;
                case 3:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Visible;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 4:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Visible;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 5:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Visible;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
                case 6:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Visible;

                    break;

                case 7:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Hidden;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Visible;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;

                case 8:
                    accueil.Visibility = Visibility.Hidden;
                    Demande.Visibility = Visibility.Hidden;
                    Traitement.Visibility = Visibility.Hidden;
                    Virement.Visibility = Visibility.Hidden;
                    Ajouter_Data.Visibility = Visibility.Visible;
                    Statistiques.Visibility = Visibility.Hidden;
                    Batch.Visibility = Visibility.Hidden;
                    Modif_Data.Visibility = Visibility.Hidden;
                    MonCompte.Visibility = Visibility.Hidden;
                    CreerSimpleUser.Visibility = Visibility.Hidden;
                    SuppSimpleUser.Visibility = Visibility.Hidden;
                    ModifSimpleUser.Visibility = Visibility.Hidden;
                    ExcelImport.Visibility = Visibility.Hidden;

                    break;
            }
        }





        // Partie Formulaire 


        private void Type_formulaire_Initialized(object sender, EventArgs e)
        {
            Type_formulaire.Items.Add("Décès Employée");
            Type_formulaire.Items.Add("Don");
            Type_formulaire.Items.Add("Autres");
        }


        private void remplir_formulaire_Click(object sender, RoutedEventArgs e)
        {
            if (Type_formulaire.SelectedItem == null || Num_demande_formulaire.Text == "_____")
            {
                MessageBox.Show("Veuillez choisir un type de prime et/ou donner le numero de demande !");

            }

            else if ( comboBox3.SelectedIndex <= -1) MessageBox.Show("Vous devez choisir une demande !");

            else
            {
                if ((string)Type_formulaire.SelectedItem == "Autres" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Autres")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Formulaire.dotx");


                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }




                        else if (field.Code.Text.Contains("TelFonct"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT TelFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                application.Selection.TypeText("0" + getExcuteScalar(Query));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Aide"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT DésignationPrime FROM dbo.TypePrime ";
                                Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                application.Selection.TypeText(getExcuteScalar(Query));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantPrime FROM dbo.TypePrime ";
                                Query += " INNER JOIN DemandePrime ON TypePrime.CodePrime = DemandePrime.CodePrime";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\Demande_" + Num_demande_formulaire.Text + ".docx");


                    MessageBoxResult result = MessageBox.Show("Voulez vous afficher le formulaire ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            application.Visible = true;
                            
                            break;

                        case MessageBoxResult.No:
                            application.Quit();

                            break;

                    }


                            

                    
                }



                else if ((string)Type_formulaire.SelectedItem == "Don" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Don")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Don.docx");
                    

                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }




                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantDem FROM DemandePrime ";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\Demande_" + Num_demande_formulaire.Text + ".docx");

                    MessageBoxResult result = MessageBox.Show("Voulez vous afficher le formulaire ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            application.Visible = true;

                            break;

                        case MessageBoxResult.No:
                            application.Quit();

                            break;

                    }
                }

                else if (Type_formulaire.Text == "Décès Employée" && verifNumDemaAndTypePrime(Num_demande_formulaire.Text) == "Décès Employée")
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Décès.docx");


                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("NumDem"))
                        {
                            try
                            {
                                field.Select();
                                application.Selection.TypeText(Num_demande_formulaire.Text);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        else if (field.Code.Text.Contains("Date"))
                        {
                            field.Select();
                            application.Selection.TypeText(DateTime.Now.ToString("dd/MM/yyyy"));
                        }


                        else if (field.Code.Text.Contains("SitFam"))
                        {
                            field.Select();
                            Query = "SELECT SitFamFonct FROM dbo.Fonctionnaire ";
                            Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                            Query += " WHERE NumDem =" + Num_demande_formulaire.Text;

                            application.Selection.TypeText(getExcuteScalar(Query));

                        }

                        else if (field.Code.Text.Contains("NomPrenom"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT NomFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string NomFonct = getExcuteScalar(Query);

                                Query = "SELECT PrenFonct FROM dbo.Fonctionnaire ";
                                Query += " INNER JOIN DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                string PrenFonct = getExcuteScalar(Query);

                                application.Selection.TypeText(NomFonct + " " + PrenFonct);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        else if (field.Code.Text.Contains("NameLastNameDem"))
                        {
                            field.Select();

                            Query = " SELECT NomParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string NomParent = getExcuteScalar(Query);

                            Query = " SELECT PrenParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string PreomParent = getExcuteScalar(Query);

                            application.Selection.TypeText(NomParent + " " + PreomParent);
                        }


                        else if (field.Code.Text.Contains("SitDem"))
                        {
                            field.Select();

                            Query = " SELECT SitFamParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string SitDem = getExcuteScalar(Query);

                            application.Selection.TypeText(SitDem);
                        }

                        else if (field.Code.Text.Contains("LienParent"))
                        {
                            field.Select();

                            Query = " SELECT LienParent FROM DemandePrime ";
                            Query += " WHERE NumDem=" + Num_demande_formulaire.Text;

                            string LienParent = getExcuteScalar(Query);

                            application.Selection.TypeText(LienParent);
                        }


                        else if (field.Code.Text.Contains("Montant"))
                        {

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                field.Select();
                                Query = "SELECT MontantDem FROM DemandePrime ";
                                Query += " WHERE NumDem =" + Num_demande_formulaire.Text;
                                command = new SqlCommand(Query, connect);
                                ConvertisseurChiffresLettres chiffre = new ConvertisseurChiffresLettres();
                                application.Selection.TypeText(command.ExecuteScalar().ToString() + " DA    ( " + chiffre.PremiereLettreMaj(chiffre.convertion((double)command.ExecuteScalar())) + " )");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\Demande_" + Num_demande_formulaire.Text + ".docx");

                    MessageBoxResult result = MessageBox.Show("Voulez vous afficher le formulaire ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            application.Visible = true;

                            break;

                        case MessageBoxResult.No:
                            application.Quit();

                            break;

                    }
                }

                else
                {
                    MessageBox.Show("Le numéro de demande et le type de prime ne conviennent pas !");

                }


            }


        }


        private void Imprimer_formulaire_Click(object sender, RoutedEventArgs e)
        {

            if (Type_formulaire.SelectedItem == null || Num_demande_formulaire.Text == "_____")
            {
                MessageBox.Show("Veuillez choisir un type de prime et/ou donner le numero de demande !");

            }

            else if (comboBox3.SelectedIndex <= -1) MessageBox.Show("Vous devez choisir une demande !");

            else
            {
                var application = new Microsoft.Office.Interop.Word.Application();
                var FormRempli = new Microsoft.Office.Interop.Word.Document();

                FormRempli = application.Documents.Add(Template: @"" + path + "\\Documents\\Demande_" + Num_demande_formulaire + ".docx");

                MessageBoxResult result = MessageBox.Show("Voulez vous afficher le formulaire avant l'impression ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        application.Visible = true;

                        MessageBoxResult imprim = MessageBox.Show("Voulez vous imprimer le formulaire maintenant ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                        switch (imprim)
                        {
                            case MessageBoxResult.Yes:
                                FormRempli.PrintOut();
                                break;

                            case MessageBoxResult.No:
                                MessageBox.Show("Vous avez choisi de ne pas l'imprimer !");
                                break;
                        }


                        break;

                    case MessageBoxResult.No:
                        break;
                }
            }

        }





        // Partie Demande

        private void Type_demandes_Initialized(object sender, EventArgs e)
        {
            Query = "SELECT * FROM TypePrime";

            SqlCommand command = new SqlCommand(this.Query, connect);

            if (connect.State == ConnectionState.Closed) connect.Open();

            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Type_prime_demande.Items.Add((string)reader["DésignationPrime"]);
            }

            if (connect.State == ConnectionState.Open) connect.Close();


        }

        private void Type_prime_demande_DropDownClosed(object sender, EventArgs e)
        {
            if ((Type_prime_demande.Text == "Décès Parent") || (Type_prime_demande.Text == "Décès Employée"))
            {
                Tab_Décès.Visibility = Visibility.Visible;
                TabControl_Demande.SelectedItem = Tab_Décès;
            }

            if (Type_prime_demande.Text == "Don")
            {
                Tab_Don.Visibility = Visibility.Visible;
                TabControl_Demande.SelectedItem = Tab_Don;
            }
        }

        private void Num_demande_formulaire_Initialized(object sender, EventArgs e)
        {
            Num_demande_formulaire.Value = lastKey;

            if (Num_demande_formulaire.Value.Length == 3) Num_demande_formulaire.Value = "00" + Num_demande_formulaire.Value;
            else if (Num_demande_formulaire.Value.Length == 4) Num_demande_formulaire.Value = "0" + Num_demande_formulaire.Value;

        }

        private void Type_deces_Initialized(object sender, EventArgs e)
        {
            Type_deces.Items.Add("Parent");
            Type_deces.Items.Add("Employée");
        }

        private void Type_deces_DropDownClosed(object sender, EventArgs e)
        {
            if (Type_deces.Text == "Parent")
            {
                // Afficher ceux du Parent
                Date_event_deces.Visibility = Visibility.Visible;
                date_evenment_deces.Visibility = Visibility.Visible;
                num_demande_deces.Visibility = Visibility.Visible;
                Num_demande_deces.Visibility = Visibility.Visible;
                nom_fonct_demande_deces.Visibility = Visibility.Visible;
                Nom_fonct_demande_deces.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces.Visibility = Visibility.Visible;
                Prenom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Date_de_demande_deces.Visibility = Visibility.Visible;
                date_demande_deces.Visibility = Visibility.Visible;
                Ajout_demande_deces_parent.Visibility = Visibility.Visible;

                // Cacher Employee
                nom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                Nom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                Prenom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                nom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Nom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                prenom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Prenom_demandeur_deces_employee.Visibility = Visibility.Hidden;
                sit_fam_demandeur_deces_employee.Visibility = Visibility.Hidden;
                Sit_fam_demandeur_deces_employee.Visibility = Visibility.Hidden;
                lien_parenté.Visibility = Visibility.Hidden;
                Lien_parenté.Visibility = Visibility.Hidden;
                Date_event_deces_employee.Visibility = Visibility.Hidden;
                date_evenment_deces_employee.Visibility = Visibility.Hidden;
                num_demande_deces_employee.Visibility = Visibility.Hidden;
                Num_demande_deces_employee.Visibility = Visibility.Hidden;
                Date_de_demande_deces_employee.Visibility = Visibility.Hidden;
                date_demande_deces_employee.Visibility = Visibility.Hidden;
                Ajout_demande_deces_employee.Visibility = Visibility.Hidden;

            }

            if (Type_deces.Text == "Employée")
            {
                // Afficher ceux de Employee
                nom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                Nom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces_employee.Visibility = Visibility.Visible;
                Prenom_fonct_demande_deces_employee.Visibility = Visibility.Hidden;
                nom_demandeur_deces_employee.Visibility = Visibility.Visible;
                Nom_demandeur_deces_employee.Visibility = Visibility.Visible;
                prenom_demandeur_deces_employee.Visibility = Visibility.Visible;
                Prenom_demandeur_deces_employee.Visibility = Visibility.Visible;
                sit_fam_demandeur_deces_employee.Visibility = Visibility.Visible;
                Sit_fam_demandeur_deces_employee.Visibility = Visibility.Visible;
                lien_parenté.Visibility = Visibility.Visible;
                Lien_parenté.Visibility = Visibility.Visible;
                Date_event_deces_employee.Visibility = Visibility.Visible;
                date_evenment_deces_employee.Visibility = Visibility.Visible;
                num_demande_deces_employee.Visibility = Visibility.Visible;
                Num_demande_deces_employee.Visibility = Visibility.Visible;
                Date_de_demande_deces_employee.Visibility = Visibility.Visible;
                date_demande_deces_employee.Visibility = Visibility.Visible;
                Ajout_demande_deces_employee.Visibility = Visibility.Visible;


                // Cacher ceux du Parent 
                Date_event_deces.Visibility = Visibility.Hidden;
                date_evenment_deces.Visibility = Visibility.Hidden;
                num_demande_deces.Visibility = Visibility.Hidden;
                Num_demande_deces.Visibility = Visibility.Hidden;
                nom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Nom_fonct_demande_deces.Visibility = Visibility.Hidden;
                prenom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Prenom_fonct_demande_deces.Visibility = Visibility.Hidden;
                Date_de_demande_deces.Visibility = Visibility.Hidden;
                date_demande_deces.Visibility = Visibility.Hidden;
                Ajout_demande_deces_parent.Visibility = Visibility.Hidden;
            }
        }

        private void Sit_fam_demandeur_deces_employee_Initialized(object sender, EventArgs e)
        {
            Sit_fam_demandeur_deces_employee.Items.Add("Mr");
            Sit_fam_demandeur_deces_employee.Items.Add("Mme");
            Sit_fam_demandeur_deces_employee.Items.Add("Melle");
        }

        private void Lien_parenté_Initialized(object sender, EventArgs e)
        {
            Lien_parenté.Items.Add("Père");
            Lien_parenté.Items.Add("Mère");
            Lien_parenté.Items.Add("Frère");
            Lien_parenté.Items.Add("Soeur");
            Lien_parenté.Items.Add("Fils");
            Lien_parenté.Items.Add("Fille");
        }

        private void Num_demande_Initialized(object sender, EventArgs e)
        {
            Num_demande.IsEnabled = false;



            if (lastKey == "")
            {
                Num_demande.Text = "001" + DateTime.Today.ToString("yy");
            }


            else
            {
                if (Int32.Parse(lastKey) % 100 == Int32.Parse(DateTime.Today.ToString("yy")))
                {
                    Num_demande.Text = ((Int32.Parse(lastKey) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy"));
                    if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                    else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;
                }

                else if (Int32.Parse(lastKey) % 100 < Int32.Parse(DateTime.Today.ToString("yy")))
                {
                    Num_demande.Text = "001" + DateTime.Today.ToString("yy");
                }
            }



        }

        private void Ajout_demande_Click(object sender, RoutedEventArgs e)
        {
            if ((string)npfonc.Content != "")
            {

                MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment ajouter cette demande ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:

                        if ((Type_prime_demande.Text == "Décès Parent") || (Type_prime_demande.Text == "Décès Employée"))
                        {
                            TabControl_Demande.SelectedItem = Tab_Décès;
                        }

                        else if (Type_prime_demande.Text == "Don")
                        {
                            TabControl_Demande.SelectedItem = Tab_Don;
                        }

                        else
                        {


                            Query = "SELECT Matricule FROM Fonctionnaire ";
                            Query += " WHERE NomFonct ='" + Nom_fonct_demande.Text + "' AND PrenFonct='" + Prenom_fonct_demande.Text + "'";

                            string Matricule = getExcuteScalar(Query);

                            Query = "SELECT CodePrime FROM TypePrime ";
                            Query += " WHERE DésignationPrime = '" + Type_prime_demande.Text + "'";

                            string CodePrime = getExcuteScalar(Query);

                            Query = "SELECT MontantPrime FROM TypePrime";
                            Query += " WHERE DésignationPrime='" + Type_prime_demande.Text + "'";

                            string Montant = getExcuteScalar(Query);

                            Query = "SELECT CompteFonct FROM Fonctionnaire";
                            Query += " WHERE Matricule=" + Matricule;

                            string CompteFonct = getExcuteScalar(Query);

                            string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

                            string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



                            Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
                            Query += " VALUES (" + Num_demande.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + MainWindow.CodeUser + ")";


                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                command = new SqlCommand(Query, connect);
                                command.ExecuteNonQuery();
                                MessageBox.Show("La demande a été ajouté !");

                                Upload();
                                Upload_traitement();








                                if (Type_prime_demande.Text == "Retraite")
                                {
                                    Query = "UPDATE Fonctionnaire SET DateDepartDefi ='" + Date_de_event + "' , MotifDepartDefi='Retraite' WHERE Matricule=" + Matricule;

                                    try
                                    {
                                        if (connect.State == ConnectionState.Closed) connect.Open();
                                        command = new SqlCommand(Query, connect);
                                        command.ExecuteNonQuery();
                                        MessageBox.Show("Départ définitif pour le fonctionnaire !");
                                        if (connect.State == ConnectionState.Open) connect.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message + "\nLa base de donnée n'a pas pu être mise à jour !");
                                    }

                                }

                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
                            }

                            finally
                            {
                                Num_demande.Text = ((Int32.Parse(Num_demande.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                                if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                                else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;

                                lastKey = Num_demande.Text;

                            }

                        }


                        break;

                    case MessageBoxResult.No:
                        MessageBox.Show("Vous avez annulé l'opération !");

                        break;

                }


            }

            else
            {
                MessageBox.Show("Veuillez choisir un fonctionnaire !");
            }

        }

        private void Ajout_demande_deces_employee_Click(object sender, RoutedEventArgs e)
        {

            if ((string)prenom_fonct_demande_deces_employee.Content != "")
            {

                MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment ajouter cette demande ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:

                        Query = "SELECT Matricule FROM Fonctionnaire ";
                        Query += " WHERE NomFonct ='" + Nom_fonct_demande_deces_employee.Text + "' AND PrenFonct='" + Prenom_fonct_demande_deces_employee.Text + "'";

                        string Matricule = getExcuteScalar(Query);

                        Query = "SELECT CodePrime FROM TypePrime ";
                        Query += " WHERE DésignationPrime = 'Décès " + Type_deces.Text + "'";

                        string CodePrime = getExcuteScalar(Query);

                        Query = "SELECT MontantPrime FROM TypePrime";
                        Query += " WHERE DésignationPrime='Décès " + Type_deces.Text + "'";

                        string Montant = getExcuteScalar(Query);

                        Query = "SELECT CompteFonct FROM Fonctionnaire";
                        Query += " WHERE Matricule=" + Matricule;

                        string CompteFonct = getExcuteScalar(Query);

                        string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande_deces_employee.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

                        string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment_deces_employee.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



                        Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser, NomParent, PrenParent, LienParent, SitFamParent )";
                        Query += " VALUES (" + Num_demande_deces_employee.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + MainWindow.CodeUser + ", '" + Nom_demandeur_deces_employee.Text + "', '" + Prenom_demandeur_deces_employee.Text + "' , '" + Lien_parenté.Text + "' , '" + Sit_fam_demandeur_deces_employee.Text + "')";


                        try
                        {
                            if (connect.State == ConnectionState.Closed) connect.Open();
                            command = new SqlCommand(Query, connect);
                            command.ExecuteNonQuery();
                            MessageBox.Show("La demande a été ajouté !");

                            Upload();
                            Upload_traitement();





                            Query = "UPDATE Fonctionnaire SET DateDepartDefi ='" + Date_de_event + "' , MotifDepartDefi='Décès' WHERE Matricule=" + Matricule;

                            try
                            {
                                if (connect.State == ConnectionState.Closed) connect.Open();
                                command = new SqlCommand(Query, connect);
                                command.ExecuteNonQuery();
                                MessageBox.Show("Départ définitif pour le fonctionnaire !");
                                if (connect.State == ConnectionState.Open) connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message + "\nLa base de donnée n'a pas pu être mise à jour !");
                            }


                            if (connect.State == ConnectionState.Open) connect.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
                        }

                        finally
                        {
                            Num_demande.Text = ((Int32.Parse(Num_demande_deces_employee.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                            if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                            else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;

                            lastKey = Num_demande_deces_employee.Text;

                        }
                        break;

                    case MessageBoxResult.No:
                        MessageBox.Show("Vous avez annulé l'opération !");

                        break;

                }

            }

            else
            {
                MessageBox.Show("Veuillez choisir un fonctionnaire !");
            }



        }

        private void Ajout_demande_deces_parent_Click(object sender, RoutedEventArgs e)
        {
            if ((string)prenom_fonct_demande_deces.Content != "")
            {
                MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment ajouter cette demande ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:

                        Query = "SELECT Matricule FROM Fonctionnaire ";
                        Query += " WHERE NomFonct ='" + Nom_fonct_demande_deces.Text + "' AND PrenFonct='" + Prenom_fonct_demande_deces.Text + "'";

                        string Matricule = getExcuteScalar(Query);

                        Query = "SELECT CodePrime FROM TypePrime ";
                        Query += " WHERE DésignationPrime = 'Décès " + Type_deces.Text + "'";

                        string CodePrime = getExcuteScalar(Query);

                        Query = "SELECT MontantPrime FROM TypePrime";
                        Query += " WHERE DésignationPrime='Décès " + Type_deces.Text + "'";

                        string Montant = getExcuteScalar(Query);

                        Query = "SELECT CompteFonct FROM Fonctionnaire";
                        Query += " WHERE Matricule=" + Matricule;

                        string CompteFonct = getExcuteScalar(Query);

                        string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

                        string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



                        Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
                        Query += " VALUES (" + Num_demande_deces.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + MainWindow.CodeUser + ")";


                        try
                        {
                            if (connect.State == ConnectionState.Closed) connect.Open();
                            command = new SqlCommand(Query, connect);
                            command.ExecuteNonQuery();
                            MessageBox.Show("La demande a été ajouté !");


                            Upload();
                            Upload_traitement();



                            if (connect.State == ConnectionState.Open) connect.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
                        }

                        finally
                        {
                            Num_demande.Text = ((Int32.Parse(Num_demande_deces.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                            if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                            else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;

                            lastKey = Num_demande_deces.Text;

                        }

                        break;

                    case MessageBoxResult.No:
                        MessageBox.Show("Vous avez annulé l'opération !");

                        break;
                }


            }
            else
            {
                MessageBox.Show("Veuillez choisir un fonctionnaire !");
            }

        }

        private void Ajout_demande_Don_Click(object sender, RoutedEventArgs e)
        {
            if ((string)prenom_fonct_demande_Don.Content != "")
            {
                MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment ajouter cette demande ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:

                        Query = "SELECT Matricule FROM Fonctionnaire ";
                        Query += " WHERE NomFonct ='" + Nom_fonct_demande_Don.Text + "' AND PrenFonct='" + Prenom_fonct_demande_Don.Text + "'";

                        string Matricule = getExcuteScalar(Query);

                        Query = "SELECT CodePrime FROM TypePrime ";
                        Query += " WHERE DésignationPrime = '" + Type_prime_demande.Text + "'";

                        string CodePrime = getExcuteScalar(Query);

                        Query = "SELECT CompteFonct FROM Fonctionnaire";
                        Query += " WHERE Matricule=" + Matricule;

                        string CompteFonct = getExcuteScalar(Query);

                        string Date_de_demande = date_demande.Text.Substring(6, 4) + "-" + date_demande.Text.Substring(3, 2) + "-" + date_demande.Text.Substring(0, 2);

                        string Date_de_event = date_evenment.Text.Substring(6, 4) + "-" + date_evenment.Text.Substring(3, 2) + "-" + date_evenment.Text.Substring(0, 2);



                        Query = "INSERT INTO DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem, DateEven, DateCreatDem, CodeUser)";
                        Query += " VALUES (" + Num_demande_Don.Text + ",'" + Date_de_demande + "'," + Matricule + "," + CodePrime + "," + Montant_don.Value + ",'" + CompteFonct + "','" + Date_de_event + "',GETDATE()," + MainWindow.CodeUser + ")";


                        try
                        {
                            if (connect.State == ConnectionState.Closed) connect.Open();
                            command = new SqlCommand(Query, connect);
                            command.ExecuteNonQuery();
                            MessageBox.Show("La demande a été ajouté !");

                            Upload();
                            Upload_traitement();






                            if (connect.State == ConnectionState.Open) connect.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + "\nLa demande n'a pas pu être ajouter !");
                        }

                        finally
                        {
                            Num_demande.Text = ((Int32.Parse(Num_demande_Don.Text) / 100) + 1).ToString() + Int32.Parse(DateTime.Today.ToString("yy")).ToString();
                            if (Num_demande.Text.Length == 3) Num_demande.Text = "00" + Num_demande.Text;
                            else if (Num_demande.Text.Length == 4) Num_demande.Text = "0" + Num_demande.Text;

                            lastKey = Num_demande_Don.Text;


                            

                        }

                        break;

                    case MessageBoxResult.No:
                        MessageBox.Show("Vous avez annulé l'opération !");

                        break;

                }

            }

            else
            {
                MessageBox.Show("Veuillez choisir un fonctionnaire !");
            }

        }



        // PARTIE Etat de virement

        private void EtatVir1_Click(object sender, RoutedEventArgs e)
        {
            if (N_Vir.SelectedIndex > -1)
            {


                if (NumCheque.Text == "_______") MessageBox.Show("Veuillez donner le numéro de chèque !");

                else
                {
                    var application = new Microsoft.Office.Interop.Word.Application();
                    var Formulaire = new Microsoft.Office.Interop.Word.Document();

                    Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Etat.docx");



                    foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                    {
                        if (field.Code.Text.Contains("Ministere"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT Ministere FROM Parametres";
                                application.Selection.TypeText(getExcuteScalar(this.Query));

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }


                        }

                        else if (field.Code.Text.Contains("Organisme"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT Organisme FROM Parametres";
                                application.Selection.TypeText(getExcuteScalar(this.Query));

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        else if (field.Code.Text.Contains("CmptSoc"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT CompteSocEsi FROM Parametres";
                                application.Selection.TypeText(getExcuteScalar(this.Query));

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }


                        }

                        else if (field.Code.Text.Contains("CmptTresor"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT CompteEsiTresor FROM Parametres";
                                application.Selection.TypeText(getExcuteScalar(this.Query));

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        else if (field.Code.Text.Contains("Somme"))
                        {
                            try
                            {

                                field.Select();

                                Query = "SELECT SUM(MontantDem) FROM DemandePrime D ";
                                Query += " INNER JOIN Virement V ON D.pv_codepv = V.pv_codepv ";
                                Query += " WHERE CodeVir = " + N_Vir.Text + " AND EtatDem='A'";

                                string somme = getExcuteScalar(this.Query);

                                StringBuilder sb = new StringBuilder();
                                if (somme.Length % 2 == 1) somme = " " + somme;
                                for (int i = 0; i < somme.Length; i++)
                                {
                                    if (i % 3 == 0)
                                        sb.Append(' ');
                                    sb.Append(somme[i]);
                                }

                                application.Selection.TypeText(sb.ToString());

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                        else if (field.Code.Text.Contains("SomeLettres"))
                        {


                            SqlCommand command = new SqlCommand(Query, connect);
                            try
                            {
                                connect.Open();
                                field.Select();
                                Query = "SELECT SUM(MontantDem) FROM DemandePrime ";
                                Query += " INNER JOIN Virement ON DemandePrime.pv_codepv = Virement.pv_codepv ";
                                Query += " WHERE CodeVir = " + N_Vir.Text + " AND EtatDem='A'";

                                ConvertisseurChiffresLettres convert = new ConvertisseurChiffresLettres();

                                double somme;



                                if (getExcuteScalar(Query) != "")
                                {
                                    somme = double.Parse(getExcuteScalar(Query));
                                }
                                else
                                {
                                    somme = 0;
                                }

                                string someLettres = convert.convertion(somme);
                                application.Selection.TypeText(someLettres.ToUpper());
                                connect.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        else if (field.Code.Text.Contains("Date"))
                        {
                            field.Select();
                            application.Selection.TypeText(date_cheque.Text);
                        }

                        else if (field.Code.Text.Contains("NumCheque"))
                        {
                            field.Select();
                            application.Selection.TypeText(NumCheque.Text);
                        }

                        else if (field.Code.Text.Contains("Observation"))
                        {
                            try
                            {

                                field.Select();
                                Query = "SELECT ObserVir FROM Virement WHERE CodeVir=" + N_Vir.Text;
                                application.Selection.TypeText(getExcuteScalar(this.Query).ToUpper());

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                        else if (field.Code.Text.Contains("NumVir"))
                        {
                            try
                            {
                                if (N_Vir.Text != null || N_Vir.Text != "")
                                {

                                    field.Select();
                                    string NumVirement = N_Vir.Text;
                                    application.Selection.TypeText(NumVirement);


                                }
                                else
                                {
                                    MessageBox.Show("Veuillez choisir un virement !");
                                }

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }


                    Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\EtatVir_" + N_Vir.Text + ".docx");

                    MessageBoxResult result = MessageBox.Show("Voulez vous afficher l'état de virement N°" + N_Vir.Text + " ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            application.Visible = true;

                            break;

                        case MessageBoxResult.No:
                            application.Quit();

                            break;
                    }
                }

                

            }

            else
            {
                MessageBox.Show("Vous devez choisir un virement !");
            }
        }


        // PARTIE AVIS & ORDRE


        public void remplirAvis (DataRow row)
        {

            if ( N_Vir.SelectedIndex > -1)
            {
                var application = new Microsoft.Office.Interop.Word.Application();
                var Formulaire = new Microsoft.Office.Interop.Word.Document();

                Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Avis.docx");


                foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                {


                    if (field.Code.Text.Contains("Name"))
                    {

                        field.Select();
                        application.Selection.TypeText((string)row["NomFonct"]);

                    }


                    else if (field.Code.Text.Contains("PRENOM"))
                    {

                        field.Select();
                        application.Selection.TypeText((string)row["PrenFonct"]);

                    }


                    else if (field.Code.Text.Contains("CmpTres"))
                    {

                        field.Select();
                        Query = "SELECT CompteEsiTresor FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(this.Query).Substring(0, 7));

                    }


                    else if (field.Code.Text.Contains("Montant"))
                    {

                        field.Select();

                        Query = " SELECT Matricule FROM Fonctionnaire WHERE NomFonct='" + (string)row["NomFonct"] + "' AND PrenFonct='" + (string)row["PrenFonct"] + "'";

                        string Matricule = getExcuteScalar(Query);

                        Query = "SELECT MontantDem FROM DemandePrime D ";
                        Query += " INNER JOIN Fonctionnaire F ON D.Matricule = F.Matricule ";
                        Query += " INNER JOIN Virement V ON D.pv_codepv = V.pv_codepv";
                        Query += " WHERE F.Matricule = " + Matricule + " AND CodeVir = " + N_Vir.Text;

                        string Total = getExcuteScalar(Query);

                        // Module pour mettre un espace entre chaque 3 chiffres

                        StringBuilder sb = new StringBuilder();
                        if (Total.Length % 2 == 1) Total = " " + Total;
                        for (int i = 0; i < Total.Length; i++)
                        {
                            if (i % 3 == 0)
                                sb.Append(' ');
                            sb.Append(Total[i]);
                        }

                        application.Selection.TypeText(sb.ToString());

                    }


                    else if (field.Code.Text.Contains("Clé"))
                    {

                        field.Select();
                        Query = "SELECT CompteEsiTresor FROM Parametres";
                        application.Selection.TypeText(getExcuteScalar(Query).Substring(8, 2));

                    }

                    else if (field.Code.Text.Contains("Cle"))
                    {

                        field.Select();
                        Query = "SELECT CompteFonct FROM Fonctionnaire WHERE Matricule = 1";
                        application.Selection.TypeText(getExcuteScalar(this.Query).Substring(18, 2));

                    }


                    else if (field.Code.Text.Contains("CompteFonct"))
                    {

                        field.Select();
                        Query = "SELECT CompteFonct FROM Fonctionnaire WHERE NomFonct='" + (string)row["NomFonct"] + "' AND PrenFonct='" + (string)row["PrenFonct"] + "'";
                        application.Selection.TypeText(getExcuteScalar(this.Query).Substring(0, 18));


                    }


                    else if (field.Code.Text.Contains("Motif"))
                    {

                        field.Select();
                        Query = "SELECT ObserVir FROM Virement Where CodeVir=" + N_Vir.Text;
                        application.Selection.TypeText(getExcuteScalar(this.Query));

                    }


                    else if (field.Code.Text.Contains("Date"))
                    {

                        field.Select();
                        application.Selection.TypeText(Date_avis.Text);
                    }
                }

                try
                {
                    Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\Avis_Ordre_" + (string)row["Matricule"] + "_" + (string)row["NomFonct"] + "_" + (string)row["PrenFonct"] + "_" + (int)row["CodePrime"] + ".docx");

                    MessageBoxResult result = MessageBox.Show("Voulez vous afficher l'avis pour " + (string)row["SitFamFonct"] + " " + (string)row["NomFonct"] + "  " + (string)row["PrenFonct"] + " ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:
                            application.Visible = true;

                            break;

                        case MessageBoxResult.No:
                            application.Quit();

                            break;

                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    MessageBox.Show(ex.Message + "\nle fichier existe déjà ou est utilisé par un autre programme !");
                }
            }


            else
            {
                MessageBox.Show("Vous devez choisir un virement !");
            }


        }



        private void Remplir_Click(object sender, RoutedEventArgs e)
        {


            if (N_Vir.SelectedIndex > -1)
            {

                Query = " SELECT * FROM DemandePrime D";
                Query += " INNER JOIN Virement V ON V.pv_codepv = D.pv_codepv";
                Query += " INNER JOIN Fonctionnaire F ON F.Matricule=D.Matricule";
                Query += " WHERE CodeVir=" + N_Vir.Text;

                command = new SqlCommand(Query, connect);

                if (connect.State == ConnectionState.Closed) { connect.Open(); }


                DataTable DT = new DataTable();

                SqlDataAdapter DA = new SqlDataAdapter(command);

                using (DA)
                {
                    DA.Fill(DT);
                }

                foreach (DataRow row in DT.Rows)
                {

                    remplirAvis(row);

                }

                MessageBox.Show("Les avis et ordres de virement ont été crée !");


                
                if (connect.State == ConnectionState.Open) { connect.Close(); }

                
                
            }
            else
            {
                MessageBox.Show("Vous devez choisir un virement !");
            }

        }


        // PARTIE LISTE


        private void RempListe_Click(object sender, RoutedEventArgs e)
        {

            if (N_Vir.SelectedIndex > -1)
            {

                var application = new Microsoft.Office.Interop.Word.Application();
                var Formulaire = new Microsoft.Office.Interop.Word.Document();

                Formulaire = application.Documents.Add(Template: @"" + path + "\\Templates\\Liste.docx");

                foreach (Microsoft.Office.Interop.Word.Field field in Formulaire.Fields)
                {

                    if (connect.State == ConnectionState.Closed) { connect.Open(); }

                    int NumPV = codePV(int.Parse(N_Vir.Text));

                    Query = "SELECT NomFonct,PrenFonct,MontantPrime,DésignationPrime,CompteFonct FROM DemandePrime ";
                    Query += " INNER JOIN Fonctionnaire ON Fonctionnaire.Matricule=DemandePrime.Matricule";
                    Query += " INNER JOIN TypePrime ON DemandePrime.CodePrime=TypePrime.CodePrime";
                    Query += " WHERE pv_codepv=" + NumPV + " AND EtatDem='A'  ORDER BY NomFonct";

                    command = new SqlCommand(Query, connect);

                    SqlDataReader reader = command.ExecuteReader();


                    if (reader.HasRows)
                    {




                        if (field.Code.Text.Contains("Name"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(0).ToString() + " " + reader.GetValue(1).ToString() + "\n");
                            }
                            reader.Close();
                        }

                        else if (field.Code.Text.Contains("Montant"))
                        {
                            field.Select();

                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(2).ToString() + "\n");
                            }
                            reader.Close();
                        }


                        else if (field.Code.Text.Contains("Motif"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(3).ToString() + "\n");
                            }
                            reader.Close();
                        }


                        else if (field.Code.Text.Contains("Compte"))
                        {
                            field.Select();
                            while (reader.Read())
                            {
                                application.Selection.TypeText(reader.GetValue(4).ToString() + "\n");
                            }
                            reader.Close();
                        }
                        reader.Close();



                        Formulaire.SaveAs2(FileName: @"" + path + "\\Documents\\ListeVirement_" + N_Vir.Text + ".docx");


                    }

                    else if (!reader.HasRows)
                    {
                        MessageBox.Show("Le Virement n'existe pas ou est vide !");
                    }

                    if (connect.State == ConnectionState.Open) { connect.Close(); }


                }

                MessageBoxResult result = MessageBox.Show("Voulez vous afficher la liste ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        application.Visible = true;

                        break;

                    case MessageBoxResult.No:
                        application.Quit();

                        break;

                }
            }
            else
            {
                MessageBox.Show("Veuillez selectionner un virement ! ");
            }
        }



        // PARTIE STATISTIQUES


        private void Diagramme_click(object sender, RoutedEventArgs e)
        {
            Diagram.DEBUT = date_debut_satats;
            Diagram.FIN = date_fin_satats;


            Stats statistiques = new Stats();


            statistiques.WindowState = WindowState.Maximized;
            statistiques.Show();
            statistiques.sourceDiagram.chart.Visibility = Visibility.Visible;
        }

        private void Cercle_Click(object sender, RoutedEventArgs e)
        {

            Diagram.DEBUT = date_debut_satats;
            Diagram.FIN = date_fin_satats;


            Stats statistiques = new Stats();



            statistiques.WindowState = WindowState.Maximized;
            statistiques.Show();
            statistiques.sourceDiagram.Don.Visibility = Visibility.Visible;

        }






        // PARTIE EXCEL


        public static int NbTablesExcel;
        public static int NumTableCourante = 1;

        public DataTable dt = new DataTable();

        private void Previous_table_Initialized(object sender, EventArgs e)
        {
            Previous_table.IsEnabled = false;
        }

        private void Next_table_Initialized(object sender, EventArgs e)
        {
            Next_table.IsEnabled = false;
        }

        private void Import_excel_Initialized(object sender, EventArgs e)
        {
            Import_excel.IsEnabled = false;
        }


        public void excelToDataGrid(int NumTableCourante, int NumTable)
        {
            // Instancier l'application Excel

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            NbTablesExcel = excelBook.Sheets.Count;
            NumTableCourante = NumTable;
            SheetNumber.Text = NumTableCourante.ToString();

            
            try
            {
                // Initialiser le DataTable et le DataGrid pour les remplir après avec les nouvelles données

                dt.Columns.Clear();
                dt.Rows.Clear();
                dtGrid_Grid.Children.Clear();
                dtGrid = new DataGrid();
                dtGrid.VerticalAlignment = VerticalAlignment.Center;
                dtGrid.HorizontalAlignment = HorizontalAlignment.Center;
                dtGrid_Grid.Children.Add(dtGrid);


                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(NumTableCourante); ;

                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;


                // Initialiser les noms des colonnes

                try
                {
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        string strColumn = "";
                        strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        dt.Columns.Add(strColumn, typeof(string));
                    }

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\nLa première ligne est réservé pour les noms de columns !");
                }


                // Remplir le DataTable puis le DataGrid avec les données de la feuille Excel

                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                   
                    
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                       
                        
                        try
                        {
                            
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            strData += strCellData + "|";

                        }

                        // Si le type des données est différent de String , on aura des exceptions à traiter

                        catch (Exception ex)
                        {

                            // Si le type est une date
                             
                            if ((excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).NumberFormat == "m/d/yyyy" )
                            {
                                douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;

                                DateTime date = DateTime.FromOADate(douCellData);

                                strData += "'" + date.ToString("MM-yy-yyyy") + "'|";

                                
                            }

                            // Si non , le type serait double
                            else
                            {
                                douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                                strData += douCellData.ToString() + "|";


                            }

                        }

                    }

                    strData = strData.Remove(strData.Length - 1, 1);
                    
                    dt.Rows.Add(strData.Split('|'));
                }

                // Definir le DataTble en tant que source pour le DataGrid afin de voir les données

                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            // Parcourir les dossiers pour choisir le fichier à importer

            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";
            

            var browsefile = openfile.ShowDialog();

            if (browsefile == true)
            {
                txtFilePath.Text = openfile.FileName;


                excelToDataGrid(NbTablesExcel, NumTableCourante);


                Previous_table.IsEnabled = true;
                Next_table.IsEnabled = true;
                Import_excel.IsEnabled = true;

            }
        }

        private void Next_table_Click(object sender, RoutedEventArgs e)
        {

            //Si le numero de la table en cours est inférieur au nombre de pages du fichier Excel 
            //On importe
            if (NumTableCourante < NbTablesExcel)
            {

                NumTableCourante++;
                if (0 < NumTableCourante && NumTableCourante <= NbTablesExcel)
                {

                    Import_excel.IsEnabled = true;
                    excelToDataGrid(NbTablesExcel, NumTableCourante);

                }
            }
            //Si non on se retrouve à la dernière page
            else MessageBox.Show("Vous êtes à la dernière table !");


        }

        private void Previous_table_Click(object sender, RoutedEventArgs e)
        {


            // Si le numero de la table en cours est supérieure à 1 on importe
            if (NumTableCourante > 1)
            {

                NumTableCourante--;
                if (0 < NumTableCourante && NumTableCourante <= NbTablesExcel)
                {

                    Import_excel.IsEnabled = true;
                    excelToDataGrid(NbTablesExcel, NumTableCourante);

                }
            }
            // Si non on est à la premire table
            else MessageBox.Show("Vous êtes à la première table !");


        }



        private void Import_excel_Click(object sender, RoutedEventArgs e)
        {

            if (Tables_BDD.Text == "Utilisateur")
            {

                try
                {

                    foreach (DataRow row in dt.Rows)
                    {
                        if (connect.State == ConnectionState.Closed) connect.Open();


                        Query = "INSERT INTO utilisateur VALUES ( ";

                        Query += "'" + row["NomUser"] + "','" + row["PrenUser"] + "','" + row["Login"] + "','" + row["MotPasse"] + "','" + row["droit"] + "'";

                        Query += ")";

                        MessageBox.Show(Query);

                        command = new SqlCommand(Query, connect);

                        command.ExecuteNonQuery();


                        if (connect.State == ConnectionState.Open) connect.Close();
                    }


                    MessageBox.Show("Les données ont été imortés !");

                    Import_excel.IsEnabled = false;
                }



                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }


            if (Tables_BDD.Text == "Fonctionnaire")
            {


                try
                {
                    // Pour chaque ligne 
                    foreach (DataRow row in dt.Rows)
                    {
                        if (connect.State == ConnectionState.Closed) connect.Open();



                        // REQUETTE POUR REMPLI CHAQUE LIGNE AVEC CES COLUMNS SPECIFIQUEMNT , LE RESTE EST A "NULL"

                        Query = "INSERT INTO Fonctionnaire VALUES ( ";

                        Query += row["Matricule"] + ",'" + row["NomFonct"] + "','" + row["PrenFonct"] + "'," + row["DateRecrut"] + "," + row["TelFonct"] + ",'";

                        Query += row["EmailFonct"] + "','" + row["CompteFonct"] + "'," + row["CodeBanque"] + ",'" + row["SitFamFonct"] + "','" + row["NomJFilleFonct"] + "',";

                        Query += row["DateDepartDefi"] + ",'" + row["MotifDepartDefi"] + "'," + row["DateDepartTmp"] + ",'" + row["MotifDepartTmp"] + "'," + row["DateRetrTmp"] + ")";

                        
                        command = new SqlCommand(Query, connect);

                        command.ExecuteNonQuery();

                        if (connect.State == ConnectionState.Open) connect.Close();
                    }


                    MessageBox.Show("Les données ont été imortés !");

                    Import_excel.IsEnabled = false;

                }



                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n\nLES DONNES EXISTENT DEJA DANS LA BASE DE DONNEES !");
                }


            }


            else
            {
                MessageBox.Show("Vous devez choisir une table pour pouvoir importer !");
            }


        }


        private void Tables_BDD_Initialized(object sender, EventArgs e)
        {
            Tables_BDD.Items.Add("Utilisateur");
            Tables_BDD.Items.Add("Fonctionnaire");
        }





        // PARTIE MISE A JOUR DES INFORMATION DE L'UTILISATEUR





        private void Mon_Compte(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Visible;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Creer_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Visible;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Supprimer_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Visible;
            ModifSimpleUser.Visibility = Visibility.Hidden;
        }

        private void Modifier_Simple_Utilisateur(object sender, RoutedEventArgs e)
        {
            MonCompte.Visibility = Visibility.Hidden;
            CreerSimpleUser.Visibility = Visibility.Hidden;
            SuppSimpleUser.Visibility = Visibility.Hidden;
            ModifSimpleUser.Visibility = Visibility.Visible;
        }

        private void Nom_fonct_demande_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Deconnexion(object sender, RoutedEventArgs e)
        {
            LOGIN w = new LOGIN();
            w.Show();
            Window window = Window.GetWindow(this);

            window.Close();
        }



        private void CreerSimpleUtilisateur_Click(object sender, RoutedEventArgs e)
        {
            if (Nom_b.Text=="" || Prenom_b.Text=="" || NomUtilisateur_b.Text=="" || MotPass_b.Password=="")
            {
                MessageBox.Show("Veuillez remplir tous les champs !");
            }

            else
            {
                if (MotPass_b.Password == MotPass_bConf.Password)
                {
                    Query = "INSERT INTO utilisateur VALUES ";
                    Query += "('" + Nom_b.Text + "','" + Prenom_b.Text + "','" + NomUtilisateur_b.Text + "','" + MotPass_b.Password + "','U')";

                    command = new SqlCommand(Query, connect);

                    try
                    {
                        connect.Open();
                        command.ExecuteNonQuery();
                        MessageBox.Show("L'utilisateur a été créer");
                        connect.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Les mots de passe sont pas identiques, veuillez reconfirmer votre mot de passe.");
                }
            }
            

        }

        private void MotPass_bConf_LostFocus(object sender, RoutedEventArgs e)
        {
            if (MotPass_b.Password == MotPass_bConf.Password)
            {
                MotPass_b.BorderBrush = Brushes.Green;
                MotPass_bConf.BorderBrush = Brushes.Green;
                MessageErreur.Visibility = Visibility.Hidden;
                if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
                {
                    CreerSimpleUtilisateur.IsEnabled = true;
                }
            }
            else
            {
                MotPass_b.BorderBrush = Brushes.Red;
                MotPass_bConf.BorderBrush = Brushes.Red;
                MessageErreur.Visibility = Visibility.Visible;
            }

        }

        private void Nom_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }

        private void Prenom_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }

        private void NomUtilisateur_b_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((Nom_b.Text != null) && (Prenom_b.Text != null) && (MotPass_b.Password != null) && (NomUtilisateur_b.Text != null) && (MotPass_bConf.Password != null))
            {
                CreerSimpleUtilisateur.IsEnabled = true;
            }
        }




        // PARTIE CHIHEB




        private void Ajout_Fonct_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(matricule.Text))
            {
                Fonctionnaire fonc = new Fonctionnaire(int.Parse(matricule.Text), nom.Text, prenom.Text, date.SelectedDate, long.Parse(Ntel.Text), email.Text, compte.Text, code.Text, sitfam.SelectedValue.ToString());
                fonc.Add_fonctionnaire();
            }
            else
            {
                MessageBox.Show("Le champ Matricule est vide");
            }
        }







        private void Upload()
        {
            DataSet ds;
            SqlDataAdapter da;

            BD con = new BD();
            con.seConnecter();
            da = con.getDataAdapter("SELECT        Fonctionnaire.NomFonct AS Nom, Fonctionnaire.PrenFonct AS Prenom, TypePrime.DésignationPrime AS Prime "
                         + " FROM            Fonctionnaire INNER JOIN "
                        + " DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule INNER JOIN "
                         + " TypePrime ON DemandePrime.CodePrime = TypePrime.CodePrime AND DemandePrime.pv_codepv IS NULL");

            ds = new DataSet();
            da.Fill(ds, "tabel1");
            dataGrid.ItemsSource = ds.Tables["tabel1"].DefaultView;
            con.seDeconnecter();

            

        }

        private void list_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Êtes-vous sûre de vouloir créer une nouvelle liste d'attente ? ", "Avertissement", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    {
                        BD con = new BD();
                        con.seConnecter();
                        string query = "UPDATE       DemandePrime  SET   pv_codepv =1000   WHERE    pv_codepv IS NULL ";
                        con.executerRequete(query);
                        Upload();
                        Upload_traitement();
                        con.seConnecter();

                    }
                    break;
                case MessageBoxResult.No:
                    Upload();
                    break;
            }
        }

        private void Upload_traitement()
        {
            DataSet ds1;
            SqlDataAdapter da1;
            BD con = new BD();
            con.seConnecter();
            da1 = con.getDataAdapter("SELECT        DemandePrime.NumDem AS Demande, Fonctionnaire.NomFonct AS NOM , Fonctionnaire.PrenFonct AS PRENOM , TypePrime.DésignationPrime AS Prime "
                         + " FROM            Fonctionnaire INNER JOIN "
                         + " DemandePrime ON Fonctionnaire.Matricule = DemandePrime.Matricule INNER JOIN"
                         + " TypePrime ON DemandePrime.CodePrime = TypePrime.CodePrime AND DemandePrime.pv_codepv = 1000");

            ds1 = new DataSet();
            da1.Fill(ds1, "tabel");
            dataGrid1.ItemsSource = ds1.Tables["tabel"].DefaultView;
            con.seConnecter();
        }



        private void decision_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BD con = new BD();
            con.seConnecter();
            ComboBox combo = sender as ComboBox;
            DataRowView dataRow = (DataRowView)dataGrid1.SelectedItem;
            int numdem = int.Parse(dataRow.Row.ItemArray[0].ToString());
            char cellValue = ' ';
            switch (combo.SelectedValue.ToString())
            {
                case "Acceptée":
                    cellValue = 'A';
                    break;

                case "Refusée":
                    cellValue = 'R';
                    break;

                case "Instance":
                    cellValue = 'I';
                    break;
            }

            
            string query = "UPDATE       DemandePrime SET                EtatDem = '" + cellValue + "'     WHERE        (NumDem = " + numdem + ") ";
            con.executerRequete(query);
            con.seConnecter();
        }



        private void dataGrid1_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var editedTextbox = e.EditingElement as TextBox;

            if (editedTextbox != null)
            {
                BD con = new BD();
                con.seConnecter();
                DataRowView dataRow = (DataRowView)dataGrid1.SelectedItem;
                int numdem = int.Parse(dataRow.Row.ItemArray[0].ToString());
                string query = "UPDATE       DemandePrime SET                MotifEtat  = '" + editedTextbox.Text + "'     WHERE        (NumDem = " + numdem + ") ";
                con.executerRequete(query);
                con.seConnecter();
            }

        }




        private void PV_Click(object sender, RoutedEventArgs e)
        {
            
                int code = 0;
                Query = "SELECT  TOP (1) CodePV FROM PV  ORDER BY CodePV DESC";

                int a = DateTime.Today.Year % 100 * 1000;
                if (getExcuteScalar(Query) != null) code = int.Parse(getExcuteScalar(Query));

                

                if (code >= a)
                {
                    code++;
                }
                else
                {
                    code = a+1;
                }

                MessageBoxResult result = MessageBox.Show("Êtes-vous sûre de vouloir créer un nouveau PV ?\nLes résultats sont irreversibles ! ", "Avertissement", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        {
                            BD con = new BD();
                            PV p = new PV(code);
                            con.seConnecter();
                            p.Add_PV();
                            con.seConnecter();
                            string query = "UPDATE       DemandePrime  SET   pv_codepv = " + code + "   WHERE    pv_codepv = 1000 AND EtatDem != 'I' ";
                            con.executerRequete(query);
                            Upload_traitement();
                            con.seDeconnecter();
                    }
                        break;
                    case MessageBoxResult.No:
                        Upload();

                        break;
                }

            
            


        }




        private void button1_Click(object sender, RoutedEventArgs e)
        {
            
            int code = 0;

            Query = "SELECT  TOP (1) CodeVir FROM Virement  ORDER BY CodeVir DESC";

            
                int a = Int32.Parse(DateTime.Today.ToString("yy")) * 1000;
            
                if (getExcuteScalar(Query) != null) code = int.Parse(getExcuteScalar(Query));
            
            if (code >= a)
                {
                    code++;
                }
                else
                {
                    code = a+1;
                }

            
            MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment créer ce virement ? ", "Attention", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    {
                        BD con = new BD();
                        con.seConnecter();

                        SqlDataReader dr1 = con.getResultatRequete("SELECT * FROM Parametres");

                        if (dr1.Read() && N_PV.SelectedIndex > -1)
                        {
                           
                            try
                            {
                                if (connect.State == ConnectionState.Closed) { connect.Open(); }

                                string query = " INSERT INTO Virement "
                                 + " (CodeVir, pv_codepv, DateCreatVir, CodeUser, MinistereVir, OrganismeVir, CompteSocVir, CompteEsiVir, BenefVir, ObserVir) "
                                + " VALUES (" + code + "," + int.Parse(N_PV.Text) + ",'" + DateTime.Today.ToString("MM-dd-yyyy") + "'," + MainWindow.CodeUser + ",'" + dr1[0].ToString() + "','" + dr1[1].ToString() + "','" + dr1[5].ToString() + "','" + dr1[6].ToString() + "','" + dr1[7].ToString() + "','" + observation.Text + "')";

                                command = new SqlCommand(query, connect);

                                command.ExecuteNonQuery();

                                MessageBox.Show("Virement N°" + code + " ajouté pour le PV N°" + N_PV.Text + " avec succès !");

                                Query = " UPDATE PV SET Virement=0 WHERE CodePV=" + N_PV.Text;

                                command = new SqlCommand(Query, connect);
                                command.ExecuteNonQuery();

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            
                            finally
                            {
                                N_PV.Items.Clear();
                                if (connect.State == ConnectionState.Open) { connect.Close(); }
                            }
                            

                        }
                    }
                    break;
                case MessageBoxResult.No:
                    Upload();

                    break;
            }
        }

        private void N_PV_DropDownOpened(object sender, EventArgs e)
        {
            N_PV.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT   CodePV FROM PV WHERE Virement = 1  ORDER BY CodePV DESC");
            while (dr.Read())
            {
                N_PV.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }

        private void comboBox_DropDownOpened(object sender, EventArgs e)
        {
            N_Vir.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT  CodeVir FROM  Virement ORDER BY CodeVir DESC");

            while (dr.Read())
            {
                N_Vir.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }



        private int codePV(int codevir)
        {
            BD con = new BD();
            con.seConnecter();
            int code = 0;
            SqlDataReader dr = con.getResultatRequete("SELECT  pv_codepv  FROM  Virement   WHERE  CodeVir = " + codevir);
            if (dr.Read())
            {
                code = (int)dr[0];
            }
            con.seDeconnecter();
            return code;


        }

        private void comboBox_DropDownOpened_1(object sender, EventArgs e)
        {
            N_Vir.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT  CodeVir   FROM Virement  ORDER BY CodeVir DESC");
            while (dr.Read())
            {
                N_Vir.Items.Add((int)dr[0]);
            }
            con.seDeconnecter();
        }


        private void ModifCompte_Click(object sender, RoutedEventArgs e)
        {
            Query = "select Login from utilisateur where CodeUser ='" + MainWindow.CodeUser + "'";
            if (string.IsNullOrEmpty(Login_box.Text))
            {
                ModifCompte.IsEnabled = false;
            }
            else
            {
                ModifCompte.IsEnabled = true;
                if (Login_box.Text != getExcuteScalar(Query))
                {

                    Query = "UPDATE utilisateur SET Login='" + Login_box.Text + "' WHERE CodeUser='" + MainWindow.CodeUser + "'";
                    command = new SqlCommand(Query, connect);

                    try
                    {
                        connect.Open();
                        command.ExecuteNonQuery();
                        MessageBox.Show("Le nom d'utilisateur a été modifier avec succée");
                        connect.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }



        }
        private void ModifCompte_Click1(object sender, RoutedEventArgs e)
        {
            Query = "SELECT Login FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
            if ((string.IsNullOrEmpty(Actuel_box.Password)) || (string.IsNullOrEmpty(Nouveau_box.Password)) || (string.IsNullOrEmpty(confirmer_box.Password)))
            {
                ModifCompte.IsEnabled = false;
            }
            else
            {
                Query = "SELECT MotPasse FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
                if (Actuel_box.Password == getExcuteScalar(Query) && (Nouveau_box.Password == confirmer_box.Password))
                {
                    ModifCompte.IsEnabled = true;
                    Query = "UPDATE utilisateur SET MotPasse='" + Nouveau_box.Password + "' WHERE CodeUser='" + MainWindow.CodeUser + "'";
                    command = new SqlCommand(Query, connect);

                    try
                    {
                        connect.Open();
                        command.ExecuteNonQuery();
                        MessageBox.Show("Le nom d'utilisateur a été modifié avec succèes");
                        connect.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    ModifCompte.IsEnabled = false;
                }
            }

        }



        private void MonCompte_Initialized(object sender, EventArgs e)
        {
            Query = "SELECT Login FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
            Login_box.Text = getExcuteScalar(Query);
            Query = "SELECT NomUser FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
            Nom_bm.Text = getExcuteScalar(Query);
            Nom_bm.IsEnabled = false;
            Query = "SELECT PrenUser FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
            Prenom_bm.Text = getExcuteScalar(Query);
            Prenom_bm.IsEnabled = false;
        }

        private void confirmer_box_LostFocus(object sender, RoutedEventArgs e)
        {
            Query = "SELECT Login FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
            if ((string.IsNullOrEmpty(Actuel_box.Password)) || (string.IsNullOrEmpty(Nouveau_box.Password)) || (string.IsNullOrEmpty(confirmer_box.Password)))
            {
                ModifCompte.IsEnabled = false;
            }
            else
            {
                if (confirmer_box.Password == Nouveau_box.Password)
                {

                    Nouveau_box.BorderBrush = Brushes.Green;
                    confirmer_box.BorderBrush = Brushes.Green;
                    MessageErreur1.Visibility = Visibility.Hidden;

                }
                else
                {

                    Nouveau_box.BorderBrush =Brushes.Red;
                    confirmer_box.BorderBrush = Brushes.Red;
                    MessageErreur1.Visibility = Visibility.Visible;

                }
            }
        }

        private void Actuel_box_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((string.IsNullOrEmpty(Actuel_box.Password)))
            {
                ModifCompte.IsEnabled = false;
            }
            else
            {
                Query = "SELECT MotPasse FROM utilisateur WHERE CodeUser ='" + MainWindow.CodeUser + "'";
                if (Actuel_box.Password == getExcuteScalar(Query))
                {
                    Actuel_box.BorderBrush = Brushes.Green;
                    MessageErreur2.Visibility = Visibility.Hidden;
                }
                else
                {
                    Actuel_box.BorderBrush = Brushes.Red;
                    MessageErreur2.Visibility = Visibility.Visible;

                }

            }

        }

        private void A_Initialized(object sender, EventArgs e)
        {
            if (MainWindow.Droit == "U")
            {
                A.Visibility = Visibility.Hidden;

            }

        }

        private void B_Initialized(object sender, EventArgs e)
        {
            if (MainWindow.Droit == "U")
            {
                B.Visibility = Visibility.Hidden;

            }

        }



        // Remplire la liste des employées accéptés pourles primes batch
        public void Remp_acc()
        {
            cotis COT = new cotis();
            SqlDataReader rd;
            BD con = new BD();
            try
            {
                con.seConnecter();
                rd = COT.batch_acc();
                while (rd.Read())
                {

                    string text1 = rd.GetString(1);
                    string text2 = rd.GetString(2);
                    int text3 = rd.GetInt32(0);
                    nom_.Items.Add(text1);
                    prenom_.Items.Add(text2);
                    matricule_.Items.Add(text3);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "message");
            }

        }

        //Rmpire la liste des employée refusé pour les primes batch
        void remp_combo()
        {



        }
        //********************************//
        void Remp_ref()
        {
            cotis COT = new cotis();
            SqlDataReader rd;
            BD con = new BD();
            try
            {
                con.seConnecter();
                rd = COT.batch_ref();
                while (rd.Read())
                {

                    string text1 = rd.GetString(1);
                    string text2 = rd.GetString(2);
                    int text3 = rd.GetInt32(0);
                    nom_.Items.Add(text1);
                    prenom_.Items.Add(text2);
                    matricule_.Items.Add(text3);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "message");
            }
        }

        //*********************************//
        private void button_Click(object sender, RoutedEventArgs e)
        {
            nom_.Items.Clear();
            prenom_.Items.Clear();
            matricule_.Items.Clear();
            Remp_acc();
            Ajouter_a_liste.IsEnabled = true;
        }

        private void button12_Click(object sender, RoutedEventArgs e)
        {
            nom_.Items.Clear();
            prenom_.Items.Clear();
            matricule_.Items.Clear();
            Remp_ref();
            Ajouter_a_liste.IsEnabled = false;
        }






        //Ajouter les employées à la liste des demandes
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Êtes-vous sûre de vouloir ajouté créer le Batch ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    if (combo_box.Text == "")
                    { MessageBox.Show("Veuillez choisir une prime"); }
                    else
                    {

                        try
                        {


                            cotis COT = new cotis();
                            BD con = new BD();
                            con.seConnecter();

                            Query = " SELECT CodePV FROM PV ORDER BY CodePV DESC ";
                            string codepv = getExcuteScalar(Query);

                            SqlDataReader R = COT.batch_acc();
                            COT.ajouter_liste(R, combo_box.Text, DateTime.Today.ToString("dd/MM/yyyy"), MainWindow.CodeUser, int.Parse(codepv));

                            

                            MessageBox.Show("Les fonctionnaires ont été ajouté pour les primes périodiques !");
                            Ajouter_a_liste.IsEnabled = false;


                            

                            Num_demande.Text = (int.Parse(lastKey)+1).ToString();

                            int code = 0;
                            int b = Int32.Parse(DateTime.Today.ToString("yy")) * 1000;

                            Query = "SELECT  TOP (1) CodePV FROM PV  ORDER BY CodePV DESC";

                            if (getExcuteScalar(Query) != null) code = int.Parse(getExcuteScalar(Query));



                            if (code > b)
                            {
                                code++;
                            }
                            else
                            {
                                code = b + 1;
                            }

                            
                            PV p = new PV(code);
                            
                            p.Add_PV();
                            

                            int code1 = 0;

                            Query = "SELECT  TOP (1) CodeVir FROM Virement  ORDER BY CodeVir DESC";


                            
                            if (getExcuteScalar(Query) != null) code1 = int.Parse(getExcuteScalar(Query));

                            if (code1 >= b)
                            {
                                code1++;
                            }
                            else
                            {
                                code1 = b + 1;
                            }

                            
                            SqlDataReader dr1 = con.getResultatRequete("SELECT * FROM Parametres");

                            if (dr1.Read())
                            {

                                try
                                {

                                    string query = " INSERT INTO Virement "
                                 + " (CodeVir, pv_codepv, DateCreatVir, CodeUser, MinistereVir, OrganismeVir, CompteSocVir, CompteEsiVir, BenefVir, ObserVir) "
                                + " VALUES (" + code1 + "," + code + ",'" + DateTime.Today.ToString("MM-dd-yyyy") + "'," + MainWindow.CodeUser + ",'" + dr1[0].ToString() + "','" + dr1[1].ToString() + "','" + dr1[5].ToString() + "','" + dr1[6].ToString() + "','" + dr1[7].ToString() + "','Prime périodeique : "+ combo_box.Text + "')";


                                    if (connect.State == ConnectionState.Closed) { connect.Open(); }

                                    
                                    command = new SqlCommand(query, connect);

                                    command.ExecuteNonQuery();

                                    MessageBox.Show("Virement N°" + code1 + " ajouté pour le PV N°" + code + " avec succès !");

                                    Query = " UPDATE PV SET Virement=0 WHERE CodePV=" + code;

                                    command = new SqlCommand(Query, connect);
                                    command.ExecuteNonQuery();


                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }

                                if (connect.State == ConnectionState.Closed) { connect.Open(); }

                            }

                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }


                        
                        Upload();
                        Upload_traitement();
                    }

                    break;

                case MessageBoxResult.No:
                    MessageBox.Show("Opération annulée !");

                    break;

            }

        }

        private void combo_box_Initialized(object sender, EventArgs e)
        {
            combo_box.Items.Add("Ramadan");
            combo_box.Items.Add("Entrée Sociale");
        }



        private List<string> Nfonc = new List<string>();
        private List<string> Pfonc = new List<string>();
        List<string> fonc = new List<string>();



        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            BD con = new BD();
            con.seConnecter();
            
            Nfonc.Clear();
            resultat.Items.Clear();
            Pfonc.Clear();
            fonc.Clear();
            SqlDataReader dr = con.getResultatRequete("SELECT  NomFonct, PrenFonct  FROM Fonctionnaire   WHERE  (NomFonct LIKE '" + Search.Text + "%') OR   (PrenFonct LIKE '" + Search.Text + "%')  ORDER BY NomFonct");

            while (dr.Read())
            {
                Nfonc.Add((string)dr[0]);
                Pfonc.Add((string)dr[1]);
                fonc.Add((string)dr[0] + "  " + (string)dr[1]);
                resultat.Items.Add((string)dr[0] + "  " + (string)dr[1]);

            }

            con.seDeconnecter();
            resultat.Visibility = Visibility.Visible;
            resultat.Visibility = Visibility.Visible;
        }

        private void Search_LostFocus(object sender, RoutedEventArgs e)
        {
            resultat.Visibility = Visibility.Collapsed;
            dataGrid.Visibility = Visibility.Visible;
            Titre0.Visibility = Visibility.Visible;


        }

        private void resultat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (resultat.SelectedIndex > -1)
            {

                //AUTRES
                Nom_fonct_demande.Text = Nfonc[resultat.SelectedIndex];
                Prenom_fonct_demande.Text = Pfonc[resultat.SelectedIndex];
                npfonc.Content = fonc[resultat.SelectedIndex];

                //DON
                Nom_fonct_demande_Don.Text = Nfonc[resultat.SelectedIndex];
                Prenom_fonct_demande_Don.Text = Pfonc[resultat.SelectedIndex];
                prenom_fonct_demande_Don.Content = fonc[resultat.SelectedIndex];


                //DECES
                prenom_fonct_demande_deces.Content = fonc[resultat.SelectedIndex];
                Nom_fonct_demande_deces.Text = Nfonc[resultat.SelectedIndex];
                Prenom_fonct_demande_deces.Text = Pfonc[resultat.SelectedIndex];

                //DECES FONCTIONNAIRE
                Nom_fonct_demande_deces_employee.Text = Nfonc[resultat.SelectedIndex];
                Prenom_fonct_demande_deces_employee.Text = Pfonc[resultat.SelectedIndex];
                prenom_fonct_demande_deces_employee.Content = fonc[resultat.SelectedIndex];


                //FORMULAIRE
                //Num_demande_formulaire.Text = 


                Search.Clear();

                accueil.Visibility = Visibility.Hidden;
                Demande.Visibility = Visibility.Visible;
                Traitement.Visibility = Visibility.Hidden;
                Virement.Visibility = Visibility.Hidden;
                Ajouter_Data.Visibility = Visibility.Hidden;
                Statistiques.Visibility = Visibility.Hidden;
                Batch.Visibility = Visibility.Hidden;
                MonCompte.Visibility = Visibility.Hidden;
                CreerSimpleUser.Visibility = Visibility.Hidden;
                SuppSimpleUser.Visibility = Visibility.Hidden;
                ModifSimpleUser.Visibility = Visibility.Hidden;
                ExcelImport.Visibility = Visibility.Hidden;
                Modif_Data.Visibility = Visibility.Hidden;

                resultat.Visibility = Visibility.Collapsed;
                dataGrid.Visibility = Visibility.Visible;
                Titre0.Visibility = Visibility.Visible;
               
            }
        }



        private void Search_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            resultat.Visibility = Visibility.Collapsed;
            dataGrid.Visibility = Visibility.Visible;
            Titre0.Visibility = Visibility.Visible;
            
        }

        private List<int> numdem = new List<int>();

        private void comboBox3_DropDownOpened(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            BD con = new BD();
            con.seConnecter();
            SqlDataReader dr = con.getResultatRequete("SELECT        DemandePrime.NumDem, Fonctionnaire.NomFonct, Fonctionnaire.PrenFonct, TypePrime.DésignationPrime "
                         + " FROM            DemandePrime INNER JOIN "
                        + " Fonctionnaire ON DemandePrime.Matricule = Fonctionnaire.Matricule INNER JOIN "
                         + " TypePrime ON DemandePrime.CodePrime = TypePrime.CodePrime "
                         + " ORDER BY DemandePrime.NumDem DESC");
            while (dr.Read())
            {
                numdem.Add((int)dr[0]);
                comboBox3.Items.Add((int)dr[0] + " " + (string)dr[1] + " " + (string)dr[2] + " " + (string)dr[3]);
            }
            con.seDeconnecter();
        }



        // HICHEM


        private void show_Click(object sender, RoutedEventArgs e)
        {

            cotis COT = new cotis();
            SqlDataReader rd;
            BD con = new BD();
            if (Liste1.Text == "") { MessageBox.Show("Veuillez choisir une donnée"); }
            else
            {
                con.seConnecter();
                SqlCommand cmd = new SqlCommand("SELECT * FROM Parametres ", con.connextion());
                rd = cmd.ExecuteReader();
                rd.Read();

                switch (Liste1.Text)
                {
                    case "Ministère":
                        Text1.Text = Convert.ToString(rd[0]);
                        break;
                    case "Organisme":
                        Text1.Text = Convert.ToString(rd[1]);
                        break;
                    case "Durée deCotisation":
                        Text1.Text = Convert.ToString(rd[2]) + "jours";
                        break;
                    case "Jour du début d'année":
                        Text1.Text = Convert.ToString(rd[3]);
                        break;
                    case "Mois du début d'année":
                        Text1.Text = Convert.ToString(rd[4]);
                        break;
                    case "Compte social de l'ESI":
                        Text1.Text = Convert.ToString(rd[5]);
                        break;
                    case "Compte Tresor de l'ESI":
                        Text1.Text = Convert.ToString(rd[6]);
                        break;
                    default:

                        break;
                }
            }
        }



        private void modif_Click(object sender, RoutedEventArgs e)
        {
            modifier();
        }


        void modifier()
        {

            MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment modifier cette donnée ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:

                    ModifierParametres mod = new ModifierParametres();
                    if (Liste1.Text == "") { MessageBox.Show("Veuillez choisir une donnée"); }
                    else
                    {
                        if (text5.Text == "") { MessageBox.Show("Veuillez entrer votre modification"); }

                        else
                        {


                            //******************
                            switch (Liste1.Text)
                            {
                                case "Ministère":
                                    mod.modifier_ministrere(text5.Text);
                                    break;
                                case "Organisme":
                                    mod.modifier_organisme(text5.Text);
                                    break;
                                case "Durée deCotisation":
                                    try { mod.modifier_cotisation(Convert.ToInt32(text5.Text)); }
                                    catch (Exception e) { MessageBox.Show("Attention la sécie doit contenir que des chiffres!!"); }

                                    break;
                                case "Jour du début d'année":
                                    try
                                    {
                                        mod.modifier_jour(Convert.ToInt32(text5.Text));
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show("Attention la sécie doit contenir que des chiffres!!");
                                    }
                                    break;
                                case "Mois du début d'année":
                                    try { mod.modifier_mois(Convert.ToInt32(text5.Text)); }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show("Attention la sécie doit contenir que des chiffres!!");
                                    }

                                    break;
                                case "Compte social de l'ESI":
                                    mod.modifier_CompteSocEsi(text5.Text);
                                    break;
                                case "Compte Tresor de l'ESI":
                                    mod.modifier_CompteEsiTresor(text5.Text);
                                    break;
                                default:

                                    break;
                            }

                        }
                    }

                    break;

                case MessageBoxResult.No:

                    MessageBox.Show("Opération annulée !");

                    break;

            }

        }


            
        


        private void MenuItem_Initialized(object sender, EventArgs e)
        {
            if (MainWindow.Droit == "U")
            {
                D.Visibility = Visibility.Collapsed;
            }
            else
            {
                D.Visibility = Visibility.Visible;
            }
        }

        private void E_Initialized(object sender, EventArgs e)
        {
            if (MainWindow.Droit == "U")
            {
                E.Visibility = Visibility.Collapsed;
            }
            else
            {
                E.Visibility = Visibility.Visible;
            }
        }

        private void F_Initialized(object sender, EventArgs e)
        {
            if (MainWindow.Droit == "U")
            {
                F.Visibility = Visibility.Collapsed;
            }
            else
            {
                F.Visibility = Visibility.Visible;
            }
        }

        private void Nom_User_Initialized(object sender, EventArgs e)
        {
            Query = "SELECT PrenUser FROM utilisateur WHERE CodeUser=" + MainWindow.CodeUser;
            Nom_User.Header = "      "+ getExcuteScalar(Query);
            Query = "SELECT NomUser FROM utilisateur WHERE CodeUser = " + MainWindow.CodeUser;
            Nom_User.Header += " " + getExcuteScalar(Query);
        }

        private void Tile_Click(object sender, RoutedEventArgs e)
        {

            System.Diagnostics.Process.Start(@"" + path + "\\aide_en_ligne\\templates\\admin\\help.html");

        }





        private void Ajout_prime_Click(object sender, RoutedEventArgs e)
        {

            cotis COT = new cotis();
            BD con = new BD();
            SqlDataReader read;
            int cpt = 1;

            con.seConnecter();
            SqlCommand cmd = new SqlCommand("SELECT * FROM TypePrime  ", con.connextion());
            read = cmd.ExecuteReader();
            while (read.Read()) { cpt++; }
            con.seDeconnecter();



            if (Montant_prime.Text == "" ||Type_prime.Text == "") { MessageBox.Show("Veuillez remplir les daux champs avant d'ajouter"); }
            else
            {
                con.seConnecter();
                cmd = new SqlCommand("SELECT * FROM TypePrime WHERE DésignationPrime='" + Convert.ToString(Type_prime.Text) + "' ", con.connextion());
                read = cmd.ExecuteReader();
                if (read.Read()) { MessageBox.Show("Cette prime existe déjà !"); }
                else
                {
                    con.seDeconnecter();

                    try
                    {
                        double mont = Convert.ToDouble(Montant_prime.Text);

                        con.seConnecter();
                        int jour = Convert.ToInt32(DateTime.Now.Day);
                        int mois = Convert.ToInt32(DateTime.Now.Month);
                        int an = Convert.ToInt32(DateTime.Now.Year);

                        string requete = "INSERT INTO  TypePrime (CodePrime, DésignationPrime, MontantPrime, DateCreatType,CodeUser)VALUES(" + (cpt + 1) + ", '" + Convert.ToString(Type_prime.Text) + "' ," + Convert.ToInt32(Montant_prime.Text) + ",'" + an + "-" + mois + "-" + jour + "',1)";
                        cmd = new SqlCommand(requete, con.connextion());
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Prime Ajoutée");
                        

                    }
                    catch (Exception a) { MessageBox.Show("Le montant ne doit pas contenir des lettres !!!"); }




                }


            }

        }

        private void Modifier_Prime_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment modifier le montant ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:

                    cotis COT = new cotis();
                    BD con = new BD();
                    if (Liste_prime.Text == "") { MessageBox.Show("Veuillez choisir la prime à modifier"); }
                    else
                    {

                        if (Modifer_M.Text == "") { MessageBox.Show("Veuillez entrer le montant"); }
                        else
                        {
                            try
                            {
                                con.seConnecter();
                                SqlCommand cmd = new SqlCommand("UPDATE TypePrime SET MontantPrime='" + Convert.ToString(Modifer_M.Text) + "'  WHERE DésignationPrime='" + Convert.ToString(Liste_prime.Text) + "' ", con.connextion());
                                cmd.ExecuteNonQuery();
                                MessageBox.Show("modification effectué");
                                con.seDeconnecter();
                            }
                            catch (SqlException ex) { MessageBox.Show("Le montant ne doit pas contenir des lettre"); }

                        }


                    }

                    break;

                case MessageBoxResult.No:

                    MessageBox.Show("Opération annulée !");

                    break;

            }
        }

        private void Afficher_Montant_Click(object sender, RoutedEventArgs e)
        {




            cotis COT = new cotis();
            BD con = new BD();
            SqlDataReader read;
            if (Liste_prime.Text == "") { MessageBox.Show("veillez choisir la prime à afficher"); }
            else
            {


                con.seConnecter();
                SqlCommand cmd = new SqlCommand("SELECT * FROM TypePrime WHERE DésignationPrime='" + Convert.ToString(Liste_prime.Text) + "' ", con.connextion());
                read = cmd.ExecuteReader();
                read.Read();
                Afficher_M.Text = Convert.ToString(read[2] + "  DINARS");

                con.seDeconnecter();
            }
        }

        private void depart_Click(object sender, RoutedEventArgs e)
        {

            MessageBoxResult result = MessageBox.Show("Voulez-vous vraiment modifier le montant ? ", "Question ?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            switch (result)
            {
                case MessageBoxResult.Yes:


                    BD con = new BD();
                    SqlCommand cmd = new SqlCommand();
                    SqlDataReader read;
                    //vérifier que l'utilisateur a bien chosit un type de départ-retour
                    if (TypeDepart.Text == "") { MessageBox.Show("veillez choisir le type de départ avant"); }
                    else //si l'utlisateur a choisit un type depuis le combobox
                    {
                        if (Entrer_matricule.Text == "") { MessageBox.Show("veillez entrer le matricule d'employé"); }
                        else
                        {
                            con.seConnecter();
                            cmd = new SqlCommand("SELECT * FROM Fonctionnaire WHERE Matricule='" + Convert.ToString(Entrer_matricule.Text) + "' ", con.connextion());

                            read = cmd.ExecuteReader();

                            if (read.Read())
                            {
                                con.seDeconnecter();



                                switch (TypeDepart.Text)
                                {
                                    case "Depart définitif":
                                        con.seConnecter();
                                        cmd = new SqlCommand("UPDATE Fonctionnaire SET DateDepartDefi = '" + date_depart.Text + "' , MotifDepartDefi='" + Convert.ToString(Entrer_motif.Text) + "', DateRetrTmp =NULL , DateDepartTmp=NULL     WHERE Matricule=" + Convert.ToString(Entrer_matricule.Text) + "", con.connextion());
                                        cmd.ExecuteNonQuery();
                                        MessageBox.Show("Modification réussite");
                                        con.seDeconnecter();
                                        break;
                                    case "Depart tempraire":
                                        con.seConnecter();
                                        cmd = new SqlCommand("UPDATE Fonctionnaire SET DateDepartTmp = '" + date_depart.Text + "' , MotifDepartTmp='" + Convert.ToString(Entrer_motif.Text) + "' ,DateRetrTmp =NULL      WHERE Matricule=" + Convert.ToString(Entrer_matricule.Text) + "", con.connextion());
                                        cmd.ExecuteNonQuery();
                                        MessageBox.Show("Modification réussite");
                                        con.seDeconnecter();

                                        break;

                                    case "retour temporaire":
                                        con.seConnecter();
                                        cmd = new SqlCommand("UPDATE Fonctionnaire SET DateRetrTmp  = '" + Convert.ToString(date_depart.Text) + "'      WHERE Matricule=" + Convert.ToString(Entrer_matricule.Text) + "", con.connextion());
                                        cmd.ExecuteNonQuery();
                                        MessageBox.Show("Modification réussite");
                                        con.seDeconnecter();
                                        break;


                                }





                            }

                            else { MessageBox.Show("Le fonctionnaire n'éxiste pas"); }
                        }

                    }


                    break;

                case MessageBoxResult.No:

                    MessageBox.Show("Opération annulée !");

                    break;

            }


        }


        //modifier les parametres//

        //1)remplire le combobox
        void Remplire_combobox()
        {
            Liste1.Items.Add("Ministère");
            Liste1.Items.Add("Organisme");
            Liste1.Items.Add("Durée deCotisation");
            Liste1.Items.Add("Jour du début d'année");
            Liste1.Items.Add("Mois du début d'année");
            Liste1.Items.Add("Compte social de l'ESI");
            Liste1.Items.Add("Compte Tresor de l'ESI");

        }
        //*****fin****//

        //Modifier les primes//
        //1)Remplire le combobox des prime
        void Remplire_Prime()
        {
            BD con = new BD();
            SqlCommand cmd = new SqlCommand();
            SqlDataReader read;
            con.seConnecter();
            cmd = new SqlCommand("SELECT * FROM TypePrime ", con.connextion());
            read = cmd.ExecuteReader();
            while (read.Read())
            {

                Liste_prime.Items.Add(read[1]);
            }
            con.seDeconnecter();


        }



        //****Findes modifications des primes****//

        //*****Mentionner le depart ou le retour des fonctionnaire******//
        //1)Remplire le combobox des départs
        void combo_depart()
        {
            TypeDepart.Items.Add("Départ définitif");
            TypeDepart.Items.Add("Départ temporaire");
            TypeDepart.Items.Add("Retour temporaire");

        }



        private void Tile_Click_1(object sender, RoutedEventArgs e)
        {
            
            System.Diagnostics.Process.Start(@""+path+ "\\Templates\\PRJ5_EQ13_CHARABI_ELALLIA_Notice_Utilisation_VirApp_2016.pdf");

        }

        private void Liste_prime_DropDownOpened(object sender, EventArgs e)
        {
            Liste_prime.Items.Clear();
            Remplire_Prime();
        }

        private void comboBox3_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex > -1) Num_demande_formulaire.Text = numdem[comboBox3.SelectedIndex].ToString();
        }

        private void A_Propo_Click(object sender, RoutedEventArgs e)
        {
            A_Propos p = new A_Propos();
            p.Show();
        }

        private void Liste1_DropDownOpened(object sender, EventArgs e)
        {
            Liste1.Items.Clear();
            Remplire_combobox();
        }

        private void TypeDepart_DropDownOpened(object sender, EventArgs e)
        {
            TypeDepart.Items.Clear();
            combo_depart();
        }


    }
}
