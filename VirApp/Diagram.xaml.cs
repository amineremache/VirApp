using Syncfusion.UI.Xaml.Charts;
using Syncfusion.Windows.Shared;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
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

namespace VirApp
{

    public partial class Diagram : UserControl
    {
        public Diagram()
        {
            InitializeComponent();

        }


        public SqlConnection connect = new SqlConnection(WorkSpace.conString);
        public string Query;
        public SqlCommand command = new SqlCommand();


        public static DateTimeEdit DEBUT;
        public static DateTimeEdit FIN;


        private void NumericalAxis_ActualRangeChanged(object sender, Syncfusion.UI.Xaml.Charts.ActualRangeChangedEventArgs e)

        {
            var axis = sender as NumericalAxis;

            //Gets the collection of visible labels.

            var labels = axis.VisibleLabels;

            if (labels != null)

            {

                if (axis.CustomLabels.Count > 0)

                    axis.CustomLabels.Clear();

                //To add the custom labels based on the position of visible labels.

                for (int index = 0; index < labels.Count; index++)

                {

                    var axisLabel = new ChartAxisLabel()

                    {
                        Position = labels[index].Position,//To set the position where the custom label should be placed.

                        //LabelContent = index//To set the content from which labels are to be taken.
                    };
                    axis.CustomLabels.Add(axisLabel);
                }
            }
        }



        public Diagramme state;

        public void setState(Diagramme newState)
        {

            state = newState;
            if (state == Diagramme.chart)
            {
                chart.Visibility = Visibility.Visible;
                Don.Visibility = Visibility.Hidden;

            }
            if (state == Diagramme.dougnout)
            {
                chart.Visibility = Visibility.Hidden;
                Don.Visibility = Visibility.Hidden;

            }
            if (state == Diagramme.histogram)
            {
                chart.Visibility = Visibility.Hidden;
                Don.Visibility = Visibility.Hidden;
            }
            if (state == Diagramme.Don)
            {
                Don.Visibility = Visibility.Visible;
                chart.Visibility = Visibility.Hidden;
            }
        }

    }

    public enum Diagramme
    {
        chart,
        dougnout,
        histogram,
        Don
    }




    public class Model
    {
        public double nbrDemande { get; set; }

        public string typeDemande { get; set; }

    }

    public class ViewModel
    {

        public ObservableCollection<Model> Demandes { get; set; }

        public SqlDataReader getResultatRequete(String requete)
        {
            SqlConnection connect = new SqlConnection(WorkSpace.conString);
            SqlDataReader resultat;
            try
            {
                connect.Open();
                SqlCommand myCommand = new SqlCommand(requete, connect);
                resultat = myCommand.ExecuteReader();
                return resultat;
            }
            catch (SqlException e)
            {
                MessageBox.Show("Erreur !" + e.Message);
                return null;
            }
            finally
            {
                connect.Close();
            }
        }

        public string getExcuteScalar(string Query)
        {
            SqlConnection connect = new SqlConnection(WorkSpace.conString);
            SqlCommand cmd = new SqlCommand(Query, connect);
            string result = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }
            //MessageBox.Show(Query);
            if (cmd.ExecuteScalar() != null) result = cmd.ExecuteScalar().ToString();
            else result = null;
            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return result;
        }


        public int nbLigne(String table, String condition)
        {
            int cpt = 0;
            string Query = "SELECT count(*) FROM " + table + " WHERE " + condition;
            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }

        public int nbLigneInner(String table, String Inner, String condition)
        {
            int cpt = 0;
            string Query = "SELECT count(*) FROM " + table;
            Query += " INNER JOIN " + Inner + " ON " + table + ".CodePrime=" + Inner + ".CodePrime";
            Query += " WHERE " + condition;

            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }

        public string dateConvert(DateTimeEdit date)
        {
            return date.Text.Substring(6, 4) + "-" + date.Text.Substring(3, 2) + "-" + date.Text.Substring(0, 2);
        }

        public int nbLigneBetween(DateTimeEdit dateDebut, DateTimeEdit dateFin)
        {
            int cpt = 0;

            string Date_de_debut = dateConvert(dateDebut);

            string Date_de_fin = dateConvert(dateFin);

            string Query = "SELECT count(*) FROM DemandePrime D";
            Query += " INNER JOIN TypePrime P ON P.CodePrime=D.CodePrime";
            Query += " INNER JOIN PV V ON V.CodePV=D.pv_codepv";
            Query += " WHERE D.EtatDem='A' AND V.DatePV BETWEEN '" + Date_de_debut + "' AND '" + Date_de_fin + "' ";

            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }



        public int nbLigne(String table)
        {
            int cpt = 0;
            string Query = "SELECT count(*) FROM " + table;
            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }



        //public ViewModel(DateTimeEdit date_debut, DateTimeEdit date_fin)
        //{

        //    this.Demandes = new ObservableCollection<Model>();


        //      string path = System.IO.Directory.GetCurrentDirectory();
        //string conString = "Data Source=(LocalDB)\\v11.0; AttachDbFilename=\"" + path + "\\BDD\\OeuvresSociales2.mdf\";Integrated Security=True";
        //SqlConnection connect = new SqlConnection(conString);
        //    string Query = "SELECT * FROM TypePrime";
        //    SqlCommand mycommand = new SqlCommand(Query, connect);

        //    try
        //    {
        //        connect.Open();
        //        SqlDataReader reader = mycommand.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            Demandes.Add(new Model() { typeDemande = (string)reader["DésignationPrime"], nbrDemande = nbLigneInner("DemandePrime", "TypePrime", "DésignationPrime='" + (string)reader["DésignationPrime"] + "'") });

        //        }
        //        reader.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        connect.Close();
        //    }







        //    //typeDemande.Add(new Model() { typeDemande = getType(1), nbrDemande = nbLigne("DemandePrime", "CodePrime=1") });

        //}


        //public ViewModel()
        //{
        //    this.Demandes = new ObservableCollection<Model>();


        //    string conString = "Data Source=localhost;Initial Catalog=OeuvresSociales2;Integrated Security=True";
        //    SqlConnection connect = new SqlConnection(conString);
        //    string Query = "SELECT * FROM TypePrime";
        //    SqlCommand mycommand = new SqlCommand(Query, connect);

        //    try
        //    {
        //        connect.Open();
        //        SqlDataReader reader = mycommand.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            Demandes.Add(new Model() { typeDemande = (string)reader["DésignationPrime"], nbrDemande = nbLigneInner("DemandePrime", "TypePrime", "DésignationPrime='" + (string)reader["DésignationPrime"] + "'") });

        //        }
        //        reader.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        connect.Close();
        //    }
        //}


        // VERSION FINALE DES STATISTIQUES

        public ViewModel()
        {
            this.Demandes = new ObservableCollection<Model>();


            SqlConnection connect = new SqlConnection(WorkSpace.conString);

            string Query = "SELECT T.CodePrime, T.DésignationPrime, count(*) AS 'Nombre Demandes'";
            Query += " FROM DemandePrime D INNER JOIN TypePrime T";
            Query += " ON D.CodePrime=T.CodePrime";
            Query += " INNER JOIN PV P";
            Query += " ON D.pv_codepv = P.CodePV";
            Query += " WHERE D.EtatDem = 'A' AND P.DatePV BETWEEN ";
            Query += " '" + dateConvert(Diagram.DEBUT) + "' AND '" + dateConvert(Diagram.FIN) + "'";
            Query += " GROUP BY T.CodePrime,T.désignationPrime";



            SqlCommand mycommand = new SqlCommand(Query, connect);

            try
            {
                connect.Open();
                SqlDataReader reader = mycommand.ExecuteReader();

                while (reader.Read())
                {
                    Demandes.Add(new Model() { typeDemande = (string)reader["DésignationPrime"], nbrDemande = double.Parse(reader["Nombre Demandes"].ToString()) });
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connect.Close();
            }
        }



    }


    public class ViewModel_4
    {
        public ViewModel_4()
        {
            this.Expenditure = new List<Model_4>();

            double total = 0;

            try
            {


                SqlConnection connect = new SqlConnection(WorkSpace.conString);

                if (connect.State == ConnectionState.Closed) connect.Open();

                string Query = "SELECT  count(*) AS 'Nombre Demandes'";
                Query += " FROM DemandePrime D INNER JOIN TypePrime T";
                Query += " ON D.CodePrime=T.CodePrime";
                Query += " INNER JOIN PV P";
                Query += " ON D.pv_codepv = P.CodePV";
                Query += " WHERE D.EtatDem = 'A' AND P.DatePV BETWEEN ";
                Query += " '" + dateConvert(Diagram.DEBUT) + "' AND '" + dateConvert(Diagram.FIN) + "'";
                Query += " GROUP BY T.CodePrime,T.désignationPrime";




                SqlCommand mycommand = new SqlCommand(Query, connect);



                SqlDataReader reader = mycommand.ExecuteReader();

                while (reader.Read())
                {
                    total += double.Parse(reader["Nombre Demandes"].ToString());
                }
                reader.Close();


                Query = "SELECT T.CodePrime, T.DésignationPrime, count(*) AS 'Nombre Demandes'";
                Query += " FROM DemandePrime D INNER JOIN TypePrime T";
                Query += " ON D.CodePrime=T.CodePrime";
                Query += " INNER JOIN PV P";
                Query += " ON D.pv_codepv = P.CodePV";
                Query += " WHERE D.EtatDem = 'A' AND P.DatePV BETWEEN ";
                Query += " '" + dateConvert(Diagram.DEBUT) + "' AND '" + dateConvert(Diagram.FIN) + "'";
                Query += " GROUP BY T.CodePrime,T.désignationPrime";

                mycommand = new SqlCommand(Query, connect);



                reader = mycommand.ExecuteReader();

                int i = 1;

                while (reader.Read())
                {
                    Expenditure.Add(new Model_4() { Expense = getType(i), Amount = double.Parse(reader["Nombre Demandes"].ToString()) / total * 100 });
                    i++;
                }

                if (connect.State == ConnectionState.Open) connect.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




            //Expenditure.Add(new Model_4() { Expense = getType(1), Amount = nbLigne("DemandePrime", "CodePrime=1") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(2), Amount = nbLigne("DemandePrime", "CodePrime=2") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(3), Amount = nbLigne("DemandePrime", "CodePrime=3") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(4), Amount = nbLigne("DemandePrime", "CodePrime=4") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(5), Amount = nbLigne("DemandePrime", "CodePrime=5") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(6), Amount = nbLigne("DemandePrime", "CodePrime=6") / total * 100 });
            //Expenditure.Add(new Model_4() { Expense = getType(7), Amount = nbLigne("DemandePrime", "CodePrime=7") / total * 100 });

        }


        public string getType(int NumPrime)
        {
            SqlConnection connect = new SqlConnection(WorkSpace.conString);
            string Query = "SELECT DésignationPrime FROM TypePrime WHERE CodePrime=" + NumPrime.ToString();
            SqlCommand command = new SqlCommand(Query, connect);
            connect.Open();

            return command.ExecuteScalar().ToString();
        }


        public string getExcuteScalar(string Query)
        {
            SqlConnection connect = new SqlConnection(WorkSpace.conString);
            SqlCommand cmd = new SqlCommand(Query, connect);
            string result = "";

            if (connect.State == ConnectionState.Closed) { connect.Open(); }
            //MessageBox.Show(Query);
            if (cmd.ExecuteScalar() != null) result = cmd.ExecuteScalar().ToString();
            else result = null;
            if (connect.State == ConnectionState.Open) { connect.Close(); }

            return result;
        }



        public int nbLigne(String table)
        {
            int cpt = 0;
            string Query = "SELECT count(*) FROM " + table;
            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }

        public int nbLigne(String table, String condition)
        {
            int cpt = 0;
            string Query = "SELECT count(*) FROM " + table + " WHERE " + condition;
            try
            {
                cpt = Int32.Parse(getExcuteScalar(Query).ToString());
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }


        public string dateConvert(DateTimeEdit date)
        {
            return date.Text.Substring(6, 4) + "-" + date.Text.Substring(3, 2) + "-" + date.Text.Substring(0, 2);
        }


        public IList<Model_4> Expenditure
        {
            get;
            set;
        }
    }

    public class Model_4
    {
        public string Expense
        {
            get;
            set;
        }

        public double Amount
        {
            get;
            set;
        }
    }

    public class Labelconvertor_4 : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            ChartAdornment pieAdornment = value as ChartAdornment;
            return String.Format("{0} %", pieAdornment.YData);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ColorConverter_4 : IValueConverter
    {
        private SolidColorBrush ApplyLight(Color color)
        {
            return new SolidColorBrush(Color.FromArgb(color.A, (byte)(color.R * 0.9), (byte)(color.G * 0.9), (byte)(color.B * 0.9)));
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null)
            {
                ChartAdornment pieAdornment = value as ChartAdornment;
                int index = pieAdornment.Series.Adornments.IndexOf(pieAdornment);
                if (index >= 7) index = 1;
                SolidColorBrush brush = pieAdornment.Series.ColorModel.CustomBrushes[index] as SolidColorBrush;
                return ApplyLight(brush.Color);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }


    }

}
