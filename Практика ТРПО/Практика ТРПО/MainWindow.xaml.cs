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
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Практика_ТРПО
{
    public static class DataGridTextSearch
    {
        public static readonly DependencyProperty SearchValueProperty =
        DependencyProperty.RegisterAttached("SearchValue", typeof(string), typeof(DataGridTextSearch),
        new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.Inherits));
        public static string GetSearchValue(DependencyObject obj)
        { return (string)obj.GetValue(SearchValueProperty); }
        public static void SetSearchValue(DependencyObject obj, string value)
        { obj.SetValue(SearchValueProperty, value); }
        public static readonly DependencyProperty IsTextMatchProperty =
        DependencyProperty.RegisterAttached("IsTextMatch", typeof(bool), typeof(DataGridTextSearch), new UIPropertyMetadata(false));
        public static bool GetIsTextMatch(DependencyObject obj)
        { return (bool)obj.GetValue(IsTextMatchProperty); }
        public static void SetIsTextMatch(DependencyObject obj, bool value)
        { obj.SetValue(IsTextMatchProperty, value); }
    }
    public class SearchValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture) { 
            string cellText = values[0] == null ? string.Empty : values[0].ToString();string searchText = values[1] as string;
            if (!string.IsNullOrEmpty(searchText) && !string.IsNullOrEmpty(cellText))
            {
                if (cellText.ToLower().IndexOf(searchText.ToLower()) != -1)return true;
                else return false;
            }       return false;
        }
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        { return null; }
    }

    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow(int reg_id,string reg_fio)
        {
            InitializeComponent();

            DG1.CanUserAddRows = false;
            DG1.CanUserDeleteRows = false;
            DG1.CanUserResizeColumns = false;
            DG1.CanUserResizeRows = false;
            DG1.CanUserReorderColumns = false;
            DG1.IsReadOnly = true;

            DG2.CanUserAddRows = false;
            DG2.CanUserDeleteRows = false;
            DG2.CanUserResizeColumns = false;
            DG2.CanUserResizeRows = false;
            DG2.CanUserReorderColumns = false;
            DG2.IsReadOnly = true;

            DG3.CanUserAddRows = false;
            DG3.CanUserDeleteRows = false;
            DG3.CanUserResizeColumns = false;
            DG3.CanUserResizeRows = false;
            DG3.CanUserReorderColumns = false;
            DG3.IsReadOnly = true;

            

            //DP1.DisplayDateStart = DateTime.Now.AddDays(0);
            //DP1.DisplayDateEnd = DateTime.Now.AddYears(2);
            //DP2.DisplayDateStart = DateTime.Now.AddDays(0);
            //DP2.DisplayDateEnd = DateTime.Now.AddYears(2);

            id_admina_from_reg = reg_id;
            admin = reg_fio;
            LabelAdm.Content += admin;

        }
        public int id_admina_from_reg;
        public string admin;

        

        public DataTable Select(string selectSQL) // функция подключения к базе данных и обработка запросов
        {
            DataTable dataTable = new DataTable("dataBase"); // создаём таблицу в приложении
            SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");                                                 // подключаемся к базе данных
            sqlConnection.Open(); // открываем базу данных
            SqlCommand sqlCommand = sqlConnection.CreateCommand(); // создаём команду
            sqlCommand.CommandText = selectSQL; // присваиваем команде текст
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand); // создаём обработчик
            sqlDataAdapter.Fill(dataTable); // возращаем таблицу с результатом
            return dataTable;

        }
        DataTable dt_user;
        DataTable dt_user1;





        void ReverseVis(bool x)
        {
            if (x == true)
            {
                DG2.Visibility = Visibility.Visible;
                Label1.Visibility = Visibility.Visible;

            }
            else { DG2.Visibility = Visibility.Hidden; Label1.Visibility = Visibility.Hidden; }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (vt.Visibility == Visibility.Hidden) vt.Visibility = Visibility.Visible;
            else vt.Visibility = Visibility.Hidden;
            DG1.SelectedIndex = -1;
            DG2.SelectedIndex = -1;
            DG3.SelectedIndex = -1;
        }
        void CompShow(int i) // Компоненты клиента
        {
            G1.Visibility = Visibility.Hidden;
            G2.Visibility = Visibility.Hidden;
            G3.Visibility = Visibility.Hidden;
            G4.Visibility = Visibility.Hidden;
            G5.Visibility = Visibility.Hidden;
            switch (i)
            {
                case 1: G1.Visibility = Visibility.Visible; break;
                case 2: G2.Visibility = Visibility.Visible; break;
                case 3: G3.Visibility = Visibility.Visible; break;
                case 4: G4.Visibility = Visibility.Visible; break;
                case 5: G5.Visibility = Visibility.Visible; break;
            }
        }
        private void Button_Click_2(object sender, RoutedEventArgs e) // Клиенты
        {
            CompShow(1);
            vt.Visibility = Visibility.Hidden;
            But1.Content = "Клиенты";
            ReverseVis(false);
            dt_user1 = Select("SELECT ID_Клиент,Фамилия,Имя,Отчество,[Контактные данные],[Паспортные данные],Адрес FROM [dbo].[Клиент] WHERE Статус is Null");

            DG1.ItemsSource = dt_user1.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            
        }
        
        DataTable UsCost;
        private void Button_Click_3(object sender, RoutedEventArgs e) // Бронь
        {
            CompShow(2);
            vt.Visibility = Visibility.Hidden;
            But1.Content = "Бронь";
            ReverseVis(true);


            UsCost = Select("SELECT [Список оказанных услуг].[ID_Брони],(Sum(Услуга.Стоимость * [Список оказанных услуг].[Количество услуг])) as Стоимость FROM[Список оказанных услуг], Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].Статус is Null GROUP BY[Список оказанных услуг].ID_Брони");

            dt_user1 = Select("SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null");
            DG1.ItemsSource = dt_user1.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
            (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
            (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";

            Nomera();
            FullCost();
        }
        public void Nomera()
        {
            dt_user1 = Select("SELECT DISTINCT ID_Номера,Номер,Этаж,[Количество комнат] as [Кол-во комнат],Класс,Стоимость FROM Номер FULL OUTER JOIN Бронь on Номер.ID_Номера = Бронь.ID_Номер WHERE (Номер.ID_Номера is NULL OR Бронь.ID_Брони is NULL or Бронь.[Дата окончания аренды] < GETDATE()) and Номер.Статус is NULL");
            DG3.ItemsSource = dt_user1.DefaultView;
            DG3.Columns[0].Visibility = Visibility.Hidden;
            (DG3.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";
        }
        public void FullCost()
        {
            for (int i = 0; i < DG1.Items.Count; i++)
            {
                DataRowView row = (DataRowView)DG1.Items[i];
                DateTime start = (DateTime)row[4];
                DateTime finish = (DateTime)row[5];
                int day_count = (finish - start).Days + 1;
                dt_user = Select($"SELECT Номер.Стоимость FROM Номер,Бронь WHERE Номер.ID_Номера = Бронь.ID_Номер and ID_Брони = {(int)row[0]}");
                double cost = double.Parse(dt_user.Rows[0]["Стоимость"].ToString());
                double day_cost = (double)(day_count * cost);
                dt_user1 = Select($"SELECT Услуга.Стоимость,[Список оказанных услуг].[Количество услуг] FROM [Список оказанных услуг],Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].ID_Брони = {(int)row[0]} and [Список оказанных услуг].Статус is Null");
                foreach (DataRow a in dt_user1.Rows)
                {
                    day_cost += (double)((decimal)a[0] * (int)a[1]);
                }
                row[6] = day_cost;

            }
        }
        private void Button_Click_4(object sender, RoutedEventArgs e) // Номера
        {
            CompShow(3);
            vt.Visibility = Visibility.Hidden;
            But1.Content = "Номера";
            ReverseVis(false);
            dt_user1 = Select("SELECT ID_Номера,Номер,Этаж,[Количество комнат],Класс,Стоимость,Примечание FROM [dbo].[Номер] WHERE Статус is NULL");
            DG1.ItemsSource = dt_user1.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";
        }

        private void Button_Click_5(object sender, RoutedEventArgs e) // Услуги
        {
            CompShow(4);
            vt.Visibility = Visibility.Hidden;
            But1.Content = "Услуги";
            ReverseVis(false);
            dt_user1 = Select("SELECT ID_Услуги,Наименование, Стоимость, [Время предоставления] FROM [dbo].[Услуга] WHERE Статус is Null");
            DG1.ItemsSource = dt_user1.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
        }

        private void Button_Click_6(object sender, RoutedEventArgs e) // Администратор
        {
            CompShow(5);
            vt.Visibility = Visibility.Hidden;
            But1.Content = "Администратор";
            ReverseVis(false);
            dt_user1 = Select("SELECT Администратор.ID_Администратора, Фамилия, Имя, Отчество, [Дата рождения],[Контактные данные],[Паспортные данные] FROM Администратор WHERE Статус is Null");
            DG1.ItemsSource = dt_user1.DefaultView;
            (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd:MM:yyyy";
            DG1.Columns[0].Visibility = Visibility.Hidden;
        }

     
        private void DG1_SelectionChanged(object sender, SelectionChangedEventArgs e) // Вывод услуг по брони
        {
            if (Convert.ToString(But1.Content) == "Бронь")
            {
                if (DG1.SelectedIndex != -1)
                {
                    id_br_red = DG1.SelectedIndex;
                    DataRowView row = (DataRowView)DG1.Items[id_br_red];
                    id_br = (int)row["ID_Брони"];

                    dt_user = Select($"SELECT [Список оказанных услуг].ID_Списка,Услуга.Наименование,Услуга.Стоимость,[Список оказанных услуг].[Количество услуг] FROM [Список оказанных услуг],Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].ID_Брони = {id_br} and [Список оказанных услуг].Статус is Null");
                    DG2.ItemsSource = dt_user.DefaultView;
                    (DG2.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
                    DG2.Columns[0].Visibility = Visibility.Hidden;
                    foreach (DataRow dr in UsCost.Rows)
                    {
                        if ((int)dr["ID_Брони"] == id_br)
                        {
                            LabelCostAllUs.Content = dr["Стоимость"];
                            LabelCostAllUs.ContentStringFormat = "f";
                            break;
                        }
                        LabelCostAllUs.Content = "";
                    }
                }
            }
            if(Convert.ToString(But1.Content) == "Клиенты")
            {
                 try
                 {  
                    id_client_red = DG1.SelectedIndex;
                    DataRowView row = (DataRowView)DG1.Items[id_client_red];
                    id_client = (int)row[0];
                 }
                 catch {  }
            }
            if(Convert.ToString(But1.Content) == "Услуги")
            {
                try
                {
                    id_usl_red = DG1.SelectedIndex;
                    DataRowView row = (DataRowView)DG1.Items[id_usl_red];
                    id_uslug = (int)row[0];
                }
                catch { }
            }
            if(Convert.ToString(But1.Content) == "Администратор")
            {
                try
                {
                    id_adm_red = DG1.SelectedIndex;
                    DataRowView row = (DataRowView)DG1.Items[id_adm_red];
                    id_adm = (int)row[0];
                }
                catch { }
            }
            if(Convert.ToString(But1.Content) == "Номера")
            {
                try
                {
                    id_nom_red = DG1.SelectedIndex;
                    DataRowView row = (DataRowView)DG1.Items[id_nom_red];
                    id_nom = (int)row[0];
                }
                catch { }
            }
        }
        int id_client = -1;
        int id_uslug = -1;
        int id_adm = -1;
        int id_nom = -1;
        private void Button_Click(object sender, RoutedEventArgs e) // Добавить Клиента
        {
            bool good = true;
            if (TB1.Text.Length == 0) { good = false; TB1.BorderBrush = Brushes.Red; }
            if (TB2.Text.Length == 0) { good = false; TB2.BorderBrush = Brushes.Red; }
            if (TB3.Text.Length == 0) { good = false; TB3.BorderBrush = Brushes.Red; }
            if (TB4.Text.Length != 13) { good = false; TB4.BorderBrush = Brushes.Red; }
            if (TB5.Text.Length != 9) { good = false; TB5.BorderBrush = Brushes.Red; }
            if (TB6.Text.Length == 0) { good = false; TB6.BorderBrush = Brushes.Red; }
            if (good == true)
            {
                SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Insert into Клиент (Фамилия,Имя,Отчество,[Контактные данные],[Паспортные данные],Адрес)values(@su,@nm,@ln,@kd,@pd,@ad)";
                cmd.Parameters.AddWithValue("@su", TB1.Text);
                cmd.Parameters.AddWithValue("@nm", TB2.Text);
                cmd.Parameters.AddWithValue("@ln", TB3.Text);
                cmd.Parameters.AddWithValue("@kd", TB4.Text);
                cmd.Parameters.AddWithValue("@pd", TB5.Text);
                cmd.Parameters.AddWithValue("@ad", TB6.Text);
                cmd.Connection = sqlConnection;
                sqlConnection.Open();
                cmd.ExecuteNonQuery();
                sqlConnection.Close();

                dt_user1 = Select("SELECT Фамилия,Имя,Отчество,[Контактные данные],[Паспортные данные],Адрес FROM [dbo].[Клиент] WHERE Статус is Null");
                DG1.ItemsSource = dt_user1.DefaultView;
                DG1.Columns[0].Visibility = Visibility.Visible;
            }
        }
        private void B4_Click(object sender, RoutedEventArgs e) // Добавить Услугу
        {
            bool good = true;
            double res = 0;
            if (TB7.Text.Length == 0) { good = false; TB7.BorderBrush = Brushes.Red; }
            if (TB8.Text.Length == 0) { good = false; TB8.BorderBrush = Brushes.Red; }
            if(double.TryParse(TB8.Text,out res) == false) { good = false; TB8.BorderBrush = Brushes.Red; }
            if (good == true)
            {
                SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Insert into Услуга (Наименование,Стоимость,[Время предоставления])values(@us,@co,@ti)";
                cmd.Parameters.AddWithValue("@us", TB7.Text);
                cmd.Parameters.AddWithValue("@co", TB8.Text);
                cmd.Parameters.AddWithValue("@ti", TB9.Text);
                cmd.Connection = sqlConnection;
                sqlConnection.Open();
                cmd.ExecuteNonQuery();
                sqlConnection.Close();
                dt_user1 = Select("SELECT ID_Услуги,Наименование, Стоимость, [Время предоставления] FROM [dbo].[Услуга] WHERE Статус is Null");
                DG1.ItemsSource = dt_user1.DefaultView;
                DG1.Columns[0].Visibility = Visibility.Hidden;
            }
        }
      

        private void B2_Click(object sender, RoutedEventArgs e) // Добавить Администратора
        {
            bool good = true;
            DateTime dr = new DateTime();
            if (TB10.Text.Length == 0) { good = false; TB10.BorderBrush = Brushes.Red; }
            if (TB11.Text.Length == 0) { good = false; TB11.BorderBrush = Brushes.Red; }
            if (TB12.Text.Length == 0) { good = false; TB12.BorderBrush = Brushes.Red; }
            if (TB13.Text.Length != 10) { good = false; TB13.BorderBrush = Brushes.Red; }
            if (TB14.Text.Length != 13) { good = false; TB14.BorderBrush = Brushes.Red; }
            if (TB15.Text.Length != 9) { good = false; TB15.BorderBrush = Brushes.Red; }
            try { dr = DateTime.Parse(TB13.Text); } catch { good = false;TB13.BorderBrush = Brushes.Red; }
            if (good == true)
            {
                
                SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Insert into Администратор (Фамилия,Имя,Отчество,[Дата Рождения],[Контактные данные],[Паспортные данные])values(@su,@na,@fn,@dr,@kd,@pd)";
                cmd.Parameters.AddWithValue("@su", TB10.Text);
                cmd.Parameters.AddWithValue("@na", TB11.Text);
                cmd.Parameters.AddWithValue("@fn", TB12.Text);
                cmd.Parameters.AddWithValue("@dr", dr);
                cmd.Parameters.AddWithValue("@kd", TB14.Text);
                cmd.Parameters.AddWithValue("@pd", TB15.Text);
                cmd.Connection = sqlConnection;
                sqlConnection.Open();
                cmd.ExecuteNonQuery();
                sqlConnection.Close();
                dt_user1 = Select("SELECT Администратор.ID_Администратора, Фамилия, Имя, Отчество, [Дата рождения],[Контактные данные],[Паспортные данные] FROM Администратор WHERE Статус is Null");
                DG1.ItemsSource = dt_user1.DefaultView;
                (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd:MM:yyyy";
                DG1.Columns[0].Visibility = Visibility.Hidden;
            }
        }

        private void B3_Click(object sender, RoutedEventArgs e) // Добавить Номер
        {
            bool good = true;
            if(TB16.Text.Length == 0 || TB16.Text == "0") { good = false; TB16.BorderBrush = Brushes.Red; }
            if(TB17.Text.Length == 0 || TB17.Text == "0") { good = false; TB17.BorderBrush = Brushes.Red; }
            if(TB21.Text.Length == 0) { good = false; TB21.BorderBrush = Brushes.Red; }
            if (good == true)
            {
                bool com = true;
                for (int i = 0; i < DG1.Items.Count; i++)
                {
                    DataRowView row = (DataRowView)DG1.Items[i];
                    if (Convert.ToString(TB21.Text) == Convert.ToString(row["Номер"]))
                    {
                        com = false;
                    }
                }
                if (com == true)
                {
                    string nclass = "";
                    switch (CB1.SelectedIndex)
                    {
                        case 0: nclass = "Эконом"; break;
                        case 1: nclass = "Люкс"; break;
                        case 2: nclass = "През. Люкс"; break;
                    }
                    SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "Insert into Номер (Этаж, [Количество комнат],Класс, Стоимость,Примечание,Номер)values(@et,@kk,@cl,@co,@pr,@nu)";
                    cmd.Parameters.AddWithValue("@et", TB16.Text);
                    cmd.Parameters.AddWithValue("@kk", TB17.Text);
                    cmd.Parameters.AddWithValue("@cl", nclass);
                    cmd.Parameters.AddWithValue("@co", TB19.Text);
                    cmd.Parameters.AddWithValue("@pr", TB20.Text);
                    cmd.Parameters.AddWithValue("@nu", TB21.Text);
                    cmd.Connection = sqlConnection;
                    sqlConnection.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection.Close();
                    dt_user1 = Select("SELECT ID_Номера,Номер,Этаж,[Количество комнат],Класс,Стоимость,Примечание FROM [dbo].[Номер] WHERE Статус is NULL");
                    DG1.ItemsSource = dt_user1.DefaultView;
                    DG1.Columns[0].Visibility = Visibility.Hidden;
                    (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";
                }
                else { MessageBox.Show("Номер не может повторяться"); TB21.Text = ""; }
            }
        }

        private void CB1_SelectionChanged(object sender, SelectionChangedEventArgs e) // Цена за Класс номера
        {
            switch (CB1.SelectedIndex)
            {
                case 0: TB19.Text = Convert.ToString(40); break;
                case 1: TB19.Text = Convert.ToString(150); break;
                case 2: TB19.Text = Convert.ToString(400); break;
            }
        }

        private void B5_Click(object sender, RoutedEventArgs e) // Добавление Брони
        { 
            bool good = true;
            if (TB18.Text.Length == 0) { good = false; TB18.BorderBrush = Brushes.Red; }
            if (TB22.Text.Length == 0) { good = false; TB22.BorderBrush = Brushes.Red; }
            if (TB23.Text.Length == 0) { good = false; TB23.BorderBrush = Brushes.Red; }
            if (TB24.Text.Length != 13) { good = false; TB24.BorderBrush = Brushes.Red; }
            if (TB25.Text.Length != 9) { good = false; TB25.BorderBrush = Brushes.Red; }
            if (TB26.Text.Length == 0) { good = false; TB26.BorderBrush = Brushes.Red; }
            if (DP1.Text.Length == 0) { good = false; MessageBox.Show("Некорректно заполнена дата начала аренды"); }
            if (DP2.Text.Length == 0) { good = false; MessageBox.Show("Некорректно заполнена дата окончания аренды"); }
            if (DP2.SelectedDate <= DP1.SelectedDate) { good = false; MessageBox.Show("Некорректно заполнены даты начала/окончания аренды"); }
            if (good == true)
            {
                if (id_nomer == -1) { MessageBox.Show("Выберите номер для аренды"); }
                else
                {
                    dt_user = Select("SELECT * From Клиент WHERE Статус is Null");
                    bool have = false;
                    int id_if_have = -1;
                    for (int i = 0; i < dt_user.Rows.Count - 1; i++)
                    {
                        if (Convert.ToString(dt_user.Rows[i][5]) == Convert.ToString(TB25.Text))
                        {
                            have = true; id_if_have = (int)(dt_user.Rows[i][0]);
                            break;
                        }
                    }
                    if (have == true)
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "Insert into Бронь (ID_Администратора,ID_Клиент,ID_Номер,[Дата начала аренды],[Дата окончания аренды])values(@ia,@ik,@in,@dn,@dk)";
                        cmd.Parameters.AddWithValue("@ia", id_admina_from_reg);
                        cmd.Parameters.AddWithValue("@ik", id_if_have);
                        cmd.Parameters.AddWithValue("@in", id_nomer);
                        cmd.Parameters.AddWithValue("@dn", DP1.Text);
                        cmd.Parameters.AddWithValue("@dk", DP2.Text);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();

                        dt_user1 = Select("SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null");
                        DG1.ItemsSource = dt_user1.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";
                        Nomera();
                        FullCost();
                    }
                    else
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "Insert into Клиент (Фамилия,Имя,Отчество,[Контактные данные],[Паспортные данные],Адрес)values(@su,@nm,@ln,@kd,@pd,@ad)";
                        cmd.Parameters.AddWithValue("@su", TB18.Text);
                        cmd.Parameters.AddWithValue("@nm", TB22.Text);
                        cmd.Parameters.AddWithValue("@ln", TB23.Text);
                        cmd.Parameters.AddWithValue("@kd", TB24.Text);
                        cmd.Parameters.AddWithValue("@pd", TB25.Text);
                        cmd.Parameters.AddWithValue("@ad", TB26.Text);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();
                        dt_user = Select("SELECT * FROM Клиент");
                        int client_index = (int)dt_user.Rows[dt_user.Rows.Count - 1][0];

                        SqlConnection sqlConnection1 = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd1 = new SqlCommand();
                        cmd1.CommandType = CommandType.Text;
                        cmd1.CommandText = "Insert into Бронь (ID_Администратора,ID_Клиент,ID_Номер,[Дата начала аренды],[Дата окончания аренды])values(@ia,@ik,@in,@dn,@dk)";
                        cmd1.Parameters.AddWithValue("@ia", id_admina_from_reg);
                        cmd1.Parameters.AddWithValue("@ik", client_index);
                        cmd1.Parameters.AddWithValue("@in", id_nomer);
                        cmd1.Parameters.AddWithValue("@dn", DP1.Text);
                        cmd1.Parameters.AddWithValue("@dk", DP2.Text);
                        cmd1.Connection = sqlConnection1;
                        sqlConnection1.Open();
                        cmd1.ExecuteNonQuery();
                        sqlConnection1.Close();
                        dt_user1 = Select("SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null");
                        DG1.ItemsSource = dt_user1.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";

                        Nomera();
                        FullCost();
                    }
                }
            }
            FullCost();
            Nomera();
        }
        int id_nomer = -1;
        private void DG3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DG3.SelectedIndex != -1)
            {
                try
                {
                    int index = DG3.SelectedIndex;
                    DataRowView row = (DataRowView)DG3.Items[index];
                    id_nomer = (int)row["ID_Номера"];
                }
                catch { MessageBox.Show("Ошибка! Некорректно выделена строка"); DG3.SelectedIndex = -1; }
            }
        }

        private void FIO_PreviewTextInput(object sender, TextCompositionEventArgs e) // ФИО маска
        {
            TextBox s = (sender as TextBox);
            char inp = e.Text[0];
            if (!Char.IsLetter(inp))
                e.Handled = true;
            if (s.Text.Length == 1)
            {
                s.Text = s.Text.ToUpper();
                s.Select(s.Text.Length, 0);
            }
            s.BorderBrush = Brushes.White;
        }


        private void KD_PreviewTextInput(object sender, TextCompositionEventArgs e) // Контактные данные маска
        {
            TextBox s = sender as TextBox;
            char inp = e.Text[e.Text.Length - 1];
            if (inp == '+' || char.IsDigit(inp)) { e.Handled = false; }else
            { e.Handled = true; }
            s.BorderBrush = Brushes.White;
        }

        private void KD_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }
        }

        private void PD_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox s = sender as TextBox;
            char inp = e.Text[e.Text.Length - 1];
            if (char.IsLetterOrDigit(inp)) { e.Handled = false; }
            else{ e.Handled = true; }
            s.Text = s.Text.ToUpper();
            s.Select(s.Text.Length, 0);
            s.BorderBrush = Brushes.White;               
        }

        private void DR_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            char inp = e.Text[e.Text.Length - 1];
            if(inp == '.' || char.IsDigit(inp))
            {
                e.Handled = false;
            }
            else { e.Handled = true; }
            tb.BorderBrush = Brushes.White;
        }

        private void Number_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            char inp = e.Text[e.Text.Length - 1];
            if (char.IsDigit(inp)) { e.Handled = false; } else { e.Handled = true; }
            tb.BorderBrush = Brushes.White;
        }

        private void Cost_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            char inp = e.Text[e.Text.Length - 1];
            if (inp == ',' || char.IsDigit(inp))
            {
                e.Handled = false;
            }
            else { e.Handled = true; }
            tb.BorderBrush = Brushes.White;
        }

        private void AD_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox s = sender as TextBox;
            s.BorderBrush = Brushes.White;
        }

        int click_count = 0;
        int id_usl = -1;
        int id_br = -1;
        int us_count = -1;

        private void B5_Copy_Click(object sender, RoutedEventArgs e) // Добавить услугу в бронь
        {        
            
            if(DG1.SelectedIndex != -1)
            {
                click_count++;
                if (click_count == 1)
                {
                    LabelCostAllUs.Visibility = Visibility.Hidden;
                    L33_Copy1.Visibility = Visibility.Hidden;
                    TB_Count.Visibility = Visibility.Visible;
                    dt_user = Select("SELECT ID_Услуги,Наименование, Стоимость FROM Услуга WHERE Статус is Null");
                    DG2.ItemsSource = dt_user.DefaultView;
                    DG2.Columns[0].Visibility = Visibility.Hidden;
                    (DG2.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
 
                }
                if(click_count == 2)
                {
                    if(id_usl != -1)
                    {
                        bool have = false;
                        int id_sp = 0;
                        dt_user = Select($"SELECT Услуга.ID_Услуги,[Список оказанных услуг].ID_Списка,Услуга.Наименование,Услуга.Стоимость,[Список оказанных услуг].[Количество услуг] FROM [Список оказанных услуг],Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].ID_Брони = {id_br} and [Список оказанных услуг].Статус is Null");
                        foreach(DataRow a in dt_user.Rows)
                        {
                            if((int)a[0] == id_usl)
                            {
                                have = true;
                                id_sp = (int)a[1];
                                us_count = (int)a[4];
                                break;
                            }
                        }
                        if (have == true)
                        {
                            
                            try
                            {
                                SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                                SqlCommand cmd = new SqlCommand();
                                cmd.CommandText = "UPDATE [Список оказанных услуг] SET [Количество услуг] = @ku WHERE Id_Списка = @is";
                                cmd.Parameters.AddWithValue("@ku", us_count += int.Parse(TB_Count.Text));
                                cmd.Parameters.AddWithValue("@is", id_sp);
                                cmd.Connection = sqlConnection;
                                sqlConnection.Open();
                                cmd.ExecuteNonQuery();
                                sqlConnection.Close();
                            }
                            catch { MessageBox.Show("Проверьте правильность введённых данных");click_count = 0; }
                        }
                        else
                        {
                            SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                            SqlCommand cmd = new SqlCommand();
                            cmd.CommandText = "INSERT into [Список оказанных услуг](ID_Услуги,ID_Брони,[Количество услуг]) values(@iu,@ib,@ku)";
                            cmd.Parameters.AddWithValue("@iu", id_usl);
                            cmd.Parameters.AddWithValue("@ib", id_br);
                            cmd.Parameters.AddWithValue("@ku", int.Parse(TB_Count.Text));

                            cmd.Connection = sqlConnection;
                            sqlConnection.Open();
                            cmd.ExecuteNonQuery();
                            sqlConnection.Close();
                        }
                        LabelCostAllUs.Visibility = Visibility.Visible;
                        L33_Copy1.Visibility = Visibility.Visible;
                        TB_Count.Visibility = Visibility.Hidden;
                    }
                    
                    dt_user = Select($"SELECT Услуга.Наименование,Услуга.Стоимость,[Список оказанных услуг].[Количество услуг] FROM [Список оказанных услуг],Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].ID_Брони = {id_br} and [Список оказанных услуг].Статус is Null");
                    DG2.ItemsSource = dt_user.DefaultView;
                    (DG2.Columns[1] as DataGridTextColumn).Binding.StringFormat = "f";
                    double sum = 0;
                    foreach(DataRowView a in DG2.Items)
                    {
                        try
                        {
                            sum += (double)((decimal)a[1] * (int)a[2]);
                        }
                        catch 
                        {   MessageBox.Show("Ошибка введённых данных");
                            L33_Copy1.Visibility = Visibility.Visible;
                            LabelCostAllUs.Visibility = Visibility.Visible;
                            TB_Count.Visibility = Visibility.Hidden;
                        }
                    }
                    LabelCostAllUs.Content = sum;
                    FullCost();
                    click_count = 0;
                }

            }
            else
            {
                MessageBox.Show("Выберите бронь, для добавления услуги");
            }
        }

        private void DG2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (click_count == 1)
            {
                try
                {
                    DataRowView row = (DataRowView)DG2.Items[DG2.SelectedIndex];
                    id_usl = (int)row["ID_Услуги"];

                }
                catch { MessageBox.Show("Ошибка выделения"); }
            }
            
                try
                {
                    DataRowView row1 = (DataRowView)DG2.Items[DG2.SelectedIndex]; //// STOPPPPPPP
                    id_spisuslug = (int)row1["ID_Списка"];
                }
                catch { }
          
        }

        int but_red_client = 0;
        int id_client_red = -1;
        private void B_Red_Client_Click(object sender, RoutedEventArgs e) // Редактировать клиента
        {
            if(id_client_red != -1)
            {
                but_red_client++;
                if(but_red_client == 1)
                {
                    DataRowView row = (DataRowView)DG1.Items[id_client_red];
                    TB1.Text = (string)row[1];
                    TB2.Text = (string)row[2];
                    TB3.Text = (string)row[3];
                    TB4.Text = (string)row[4];
                    TB5.Text = (string)row[5];
                    TB6.Text = (string)row[6];
                }
                if(but_red_client == 2)
                {
                    bool good = true;
                    if (TB1.Text.Length == 0) { good = false; TB1.BorderBrush = Brushes.Red; }
                    if (TB2.Text.Length == 0) { good = false; TB2.BorderBrush = Brushes.Red; }
                    if (TB3.Text.Length == 0) { good = false; TB3.BorderBrush = Brushes.Red; }
                    if (TB4.Text.Length != 13) { good = false; TB4.BorderBrush = Brushes.Red; }
                    if (TB5.Text.Length != 9) { good = false; TB5.BorderBrush = Brushes.Red; }
                    if (TB6.Text.Length == 0) { good = false; TB6.BorderBrush = Brushes.Red; }
                    if (good == true)
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Клиент SET Фамилия = @1,Имя = @2, Отчество = @3,[Контактные данные] = @4,[Паспортные данные] = @5,Адрес = @6 WHERE ID_Клиент = @7";
                        cmd.Parameters.AddWithValue("@1", TB1.Text);
                        cmd.Parameters.AddWithValue("@2", TB2.Text);
                        cmd.Parameters.AddWithValue("@3", TB3.Text);
                        cmd.Parameters.AddWithValue("@4", TB4.Text);
                        cmd.Parameters.AddWithValue("@5", TB5.Text);
                        cmd.Parameters.AddWithValue("@6", TB6.Text);
                        cmd.Parameters.AddWithValue("@7", id_client);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();

                        dt_user = Select("SELECT * FROM Клиент WHERE Статус is Null");
                        DG1.ItemsSource = dt_user.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        DG1.Columns[7].Visibility = Visibility.Hidden;
                        TB1.Text = ""; TB2.Text = ""; TB3.Text = ""; TB4.Text = ""; TB5.Text = ""; TB6.Text = "";
                        but_red_client = 0;
                    }
                    else { but_red_client--; }
                }
            }
            else
            {
                MessageBox.Show("Выберите клиента для редактирования");
            }
        }

        int but_red_usl = 0;
        int id_usl_red = -1;
        private void B_Red_Usl_Click(object sender, RoutedEventArgs e)
        {
            if (id_usl_red != -1)
            {
                but_red_usl++;
                if (but_red_usl == 1)
                {
                    DataRowView row = (DataRowView)DG1.Items[id_usl_red];
                    TB7.Text = (string)row[1];
                    TB8.Text = row[2].ToString();
                    TB9.Text = row[3].ToString();
                }
                if (but_red_usl == 2)
                {
                    bool good = true;
                    double res = 0;
                    if (TB7.Text.Length == 0) { good = false; TB7.BorderBrush = Brushes.Red; }
                    if (TB8.Text.Length == 0) { good = false; TB8.BorderBrush = Brushes.Red; }
                    if (double.TryParse(TB8.Text, out res) == false) { good = false; TB8.BorderBrush = Brushes.Red; }
                    if (good == true)
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Услуга SET Наименование = @1,Стоимость = @2, [Время предоставления] = @3 WHERE ID_Услуги = @4";
                        cmd.Parameters.AddWithValue("@1", TB7.Text);
                        cmd.Parameters.AddWithValue("@2", double.Parse(TB8.Text));
                        cmd.Parameters.AddWithValue("@3", TB9.Text);
                        cmd.Parameters.AddWithValue("@4", id_uslug);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();

                        dt_user = Select("SELECT * FROM Услуга WHERE Статус is Null");
                        DG1.ItemsSource = dt_user.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
                        DG1.Columns[4].Visibility = Visibility.Hidden;

                        TB7.Text = ""; TB8.Text = ""; TB9.Text = "";
                        but_red_usl = 0;
                    }
                    else { but_red_usl--; }
                }
            }
            else
            {
                MessageBox.Show("Выберите услугу для редактирования");
            }
        }

        int but_red_adm = 0;
        int id_adm_red = -1;
        private void B_Red_Adm_Click(object sender, RoutedEventArgs e)
        {
            if (id_adm_red != -1)
            {
                but_red_adm++;
                if (but_red_adm == 1)
                {
                    DataRowView row = (DataRowView)DG1.Items[id_adm_red];
                    DateTime d = DateTime.Parse(row[4].ToString());
                    TB10.Text = row[1].ToString();
                    TB11.Text = row[2].ToString();
                    TB12.Text = row[3].ToString();
                    TB13.Text = d.ToShortDateString();
                    TB14.Text = row[5].ToString();
                    TB15.Text = row[6].ToString();
                }
                if (but_red_adm == 2)
                {
                    bool good = true;
                    DateTime dr = new DateTime();
                    if (TB10.Text.Length == 0) { good = false; TB10.BorderBrush = Brushes.Red; }
                    if (TB11.Text.Length == 0) { good = false; TB11.BorderBrush = Brushes.Red; }
                    if (TB12.Text.Length == 0) { good = false; TB12.BorderBrush = Brushes.Red; }
                    if (TB13.Text.Length != 10) { good = false; TB13.BorderBrush = Brushes.Red; }
                    if (TB14.Text.Length != 13) { good = false; TB14.BorderBrush = Brushes.Red; }
                    if (TB15.Text.Length != 9) { good = false; TB15.BorderBrush = Brushes.Red; }
                    try { dr = DateTime.Parse(TB13.Text); } catch { good = false; TB13.BorderBrush = Brushes.Red; }
                    if (good == true)
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Администратор SET Фамилия = @1,Имя = @2,Отчество = @3,[Дата Рождения] = @4,[Контактные данные] = @5,[Паспортные данные] = @6 WHERE ID_Администратора = @7";
                        cmd.Parameters.AddWithValue("@1", TB10.Text.ToString());
                        cmd.Parameters.AddWithValue("@2", TB11.Text.ToString());
                        cmd.Parameters.AddWithValue("@3", TB12.Text.ToString());
                        cmd.Parameters.AddWithValue("@4", dr);
                        cmd.Parameters.AddWithValue("@5", TB14.Text.ToString());
                        cmd.Parameters.AddWithValue("@6", TB15.Text.ToString());
                        cmd.Parameters.AddWithValue("@7", id_adm);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();

                        dt_user = Select("SELECT ID_Администратора,Фамилия,Имя,Отчество,[Дата Рождения],[Контактные данные],[Паспортные данные] FROM Администратор WHERE Статус is Null");
                        DG1.ItemsSource = dt_user.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd:MM:yyyy";

                        TB10.Text = ""; TB11.Text = ""; TB12.Text = "";TB13.Text = ""; TB14.Text = ""; TB15.Text = "";
                        but_red_adm= 0;
                    }
                    else { but_red_adm--; }
                }
            }
            else
            {
                MessageBox.Show("Выберите администратора для редактирования");
            }
        }

        int but_red_nom = 0;
        int id_nom_red = -1;
        private void B_Red_Nom_Click(object sender, RoutedEventArgs e)
        {
            if (id_nom_red != -1)
            {
                but_red_nom++;
                if (but_red_nom == 1)
                {
                    DataRowView row = (DataRowView)DG1.Items[id_nom_red];
                    TB16.Text = row[2].ToString();
                    TB17.Text = row[3].ToString();
                    if (row[4].ToString() == "Эконом") CB1.SelectedIndex = 0; 
                    if (row[4].ToString() == "Люкс") CB1.SelectedIndex = 1; 
                    if (row[4].ToString() == "През. Люкс") CB1.SelectedIndex = 2;
                    TB20.Text = row[6].ToString();
                    TB21.Text = row[1].ToString();
                }
                if (but_red_nom == 2)
                {
                    bool good = true;
                    if (TB16.Text.Length == 0 || TB16.Text == "0") { good = false; TB16.BorderBrush = Brushes.Red; }
                    if (TB17.Text.Length == 0 || TB17.Text == "0") { good = false; TB17.BorderBrush = Brushes.Red; }
                    if (TB21.Text.Length == 0) { good = false; TB21.BorderBrush = Brushes.Red; }
                    for (int i = 0; i < DG1.Items.Count; i++)
                    {
                        DataRowView row = (DataRowView)DG1.Items[i];
                        if (id_nom_red != i)
                        {
                            if (Convert.ToString(TB21.Text) == Convert.ToString(row["Номер"]))
                            {
                                good = false;
                            }
                        }
                    }
                    if (good == true)
                    {
                        string nclass = "";
                        switch (CB1.SelectedIndex)
                        {
                            case 0: nclass = "Эконом"; break;
                            case 1: nclass = "Люкс"; break;
                            case 2: nclass = "През. Люкс"; break;
                        }
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Номер SET Этаж = @1,[Количество комнат] = @2,Класс = @3,Стоимость = @4,Примечание = @5,Номер = @6 WHERE ID_Номера = @7";
                        cmd.Parameters.AddWithValue("@1", int.Parse(TB16.Text));
                        cmd.Parameters.AddWithValue("@2", int.Parse(TB17.Text));
                        cmd.Parameters.AddWithValue("@3", nclass);
                        cmd.Parameters.AddWithValue("@4", int.Parse(TB19.Text));
                        cmd.Parameters.AddWithValue("@5", TB20.Text);
                        cmd.Parameters.AddWithValue("@6", int.Parse(TB21.Text));
                        cmd.Parameters.AddWithValue("@7", id_nom);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();

                        dt_user1 = Select("SELECT ID_Номера,Номер,Этаж,[Количество комнат],Класс,Стоимость,Примечание FROM [dbo].[Номер] WHERE Статус is NULL");
                        DG1.ItemsSource = dt_user1.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";

                        TB16.Text = ""; TB17.Text = ""; TB20.Text = ""; TB21.Text = "";
                        but_red_nom = 0;
                    }
                    else { but_red_nom--; MessageBox.Show("Проверьте правильность введённых данных"); }
                }
            }
            else
            {
                MessageBox.Show("Выберите номер для редактирования");
            }
        }

        int but_red_br = 0;
        int id_br_red = -1;
        private void B_Red_Br_Click(object sender, RoutedEventArgs e)
        {
            if (id_br_red != -1)
            {
                but_red_br++;
                if (but_red_br == 1)
                {
                    DataRowView row = (DataRowView)DG1.Items[id_br_red];
                    DateTime start = DateTime.Parse(row[4].ToString());
                    DateTime finish = DateTime.Parse(row[5].ToString());
                    DP1.Text = (DateTime.Parse(row[4].ToString())).ToShortDateString();
                    DP2.Text = (DateTime.Parse(row[5].ToString())).ToShortDateString();                    
                }
                if (but_red_br == 2)
                {
                    bool good = true;
                    if (DP1.Text.Length == 0) { good = false; MessageBox.Show("Некорректно заполнена дата начала аренды"); }
                    if (DP2.Text.Length == 0) { good = false; MessageBox.Show("Некорректно заполнена дата окончания аренды"); }
                    if (DP2.SelectedDate <= DP1.SelectedDate) { good = false; MessageBox.Show("Некорректно заполнены даты начала/окончания аренды"); }
                    if (good == true)
                    {
                        if (DG3.SelectedIndex == -1)
                        {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Бронь SET [Дата начала аренды] = @1,[Дата окончания аренды] = @2 WHERE ID_Брони = @3";
                        cmd.Parameters.AddWithValue("@1", DP1.SelectedDate);
                        cmd.Parameters.AddWithValue("@2", DP2.SelectedDate);
                        cmd.Parameters.AddWithValue("@3", id_br);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();
                        }
                        else
                        {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Бронь SET ID_Номер = @0,[Дата начала аренды] = @1,[Дата окончания аренды] = @2 WHERE ID_Брони = @3";
                        cmd.Parameters.AddWithValue("@0", id_nomer);
                        cmd.Parameters.AddWithValue("@1", DP1.SelectedDate);
                        cmd.Parameters.AddWithValue("@2", DP2.SelectedDate);
                        cmd.Parameters.AddWithValue("@3", id_br);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();
                        }

                        dt_user1 = Select("SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null");
                        DG1.ItemsSource = dt_user1.DefaultView;
                        DG1.Columns[0].Visibility = Visibility.Hidden;
                        (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                        (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";

                        Nomera();
                        FullCost();

                        DP1.Text = ""; DP2.Text = "";
                        but_red_br = 0;
                    }
                    else { but_red_br--; MessageBox.Show("Проверьте правильность введённых данных"); }
                }
            }
            else
            {
                MessageBox.Show("Выберите бронь для редактирования");
            }
        }

        private void B_Del_Nom_Click(object sender, RoutedEventArgs e)
        {
            if(id_nom_red != -1)
            {
                SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "UPDATE Номер SET Статус = @1 WHERE ID_Номера = @2";
                cmd.Parameters.AddWithValue("@1", "Delete");
                cmd.Parameters.AddWithValue("@2", id_nom);
                cmd.Connection = sqlConnection;
                sqlConnection.Open();
                cmd.ExecuteNonQuery();
                sqlConnection.Close();
                dt_user = Select("SELECT Номер FROM Номер,Бронь WHERE Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null and Бронь.[Дата окончания аренды] < GETDATE()");
                foreach (DataRow a in dt_user.Rows)
                {
                    if(id_nom == (int)a["Номер"]) { MessageBox.Show("Внимание, данный номер арендован\nРекомендуется изменить запись в таблице \"Бронь\""); }
                }
                dt_user1 = Select("SELECT ID_Номера,Номер,Этаж,[Количество комнат],Класс,Стоимость,Примечание FROM [dbo].[Номер] WHERE Статус is NULL");
                DG1.ItemsSource = dt_user1.DefaultView;
                DG1.Columns[0].Visibility = Visibility.Hidden;
                (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";
            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }

        private void B_Del_Br_Click(object sender, RoutedEventArgs e)
        {
            if (id_br_red != -1)
            {
                MessageBoxResult mb = MessageBox.Show("Вы действительно хотите удалить запись?","Удаление записи",MessageBoxButton.YesNo);
                if (mb == MessageBoxResult.Yes)
                {
                    SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "UPDATE Бронь SET Статус = @1,[Дата окончания аренды] = GETDATE() WHERE ID_Брони = @2";
                    cmd.Parameters.AddWithValue("@1", "Delete");
                    cmd.Parameters.AddWithValue("@2", id_br);
                    cmd.Connection = sqlConnection;
                    sqlConnection.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection.Close();

                    dt_user1 = Select("SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null");
                    DG1.ItemsSource = dt_user1.DefaultView;
                    DG1.Columns[0].Visibility = Visibility.Hidden;
                    (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                    (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
                    (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";
                    DG2.ItemsSource = null;
                    DG2.Columns.Clear();
                    Nomera();
                    FullCost();
                    LabelCostAllUs.Content = "";
                }
            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }

        private void B_Del_Adm_Click(object sender, RoutedEventArgs e)
        {
            if (id_adm_red != -1)
            {
                MessageBoxResult mb = MessageBox.Show("Вы действительно хотите удалить запись?", "Удаление записи", MessageBoxButton.YesNo);
                if (mb == MessageBoxResult.Yes)
                {
                    if (id_adm != id_admina_from_reg)
                    {
                        SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandText = "UPDATE Администратор SET Статус = @1 WHERE ID_Администратора = @2";
                        cmd.Parameters.AddWithValue("@1", "Delete");
                        cmd.Parameters.AddWithValue("@2", id_adm);
                        cmd.Connection = sqlConnection;
                        sqlConnection.Open();
                        cmd.ExecuteNonQuery();
                        sqlConnection.Close();
                        dt_user1 = Select("SELECT Администратор.ID_Администратора, Фамилия, Имя, Отчество, [Дата рождения],[Контактные данные],[Паспортные данные] FROM Администратор WHERE Статус is Null");
                        DG1.ItemsSource = dt_user1.DefaultView;
                        (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd:MM:yyyy";
                        DG1.Columns[0].Visibility = Visibility.Hidden;

                    }
                    else { MessageBox.Show("Невозможно удалить авторизованного администратора"); }
                }

            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }

        private void B_Filter_Adm_Click(object sender, RoutedEventArgs e)
        {
            string s = "SELECT ID_Администратора, Фамилия, Имя, Отчество, [Дата Рождения],[Контактные данные],[Паспортные данные] FROM Администратор WHERE Статус is NULL";
            if(TB10.Text.Length != 0) { s += $" and Фамилия = '{TB10.Text}'"; }
            if(TB11.Text.Length != 0) { s += $" and Имя = '{TB11.Text}'"; }
            if(TB12.Text.Length != 0) { s += $" and Отчество = '{TB12.Text}'"; }
            if(TB13.Text.Length != 0) {
                DateTime dr = DateTime.Parse(TB13.Text);
                string yyyy = dr.Year.ToString(); string mm = ""; string dd = "";
                if(dr.Month < 10) { mm = "0"+dr.Month.ToString(); } else { mm = dr.Month.ToString(); }
                if(dr.Day < 10) { dd = "0" + dr.Day.ToString(); } else { dd = dr.Day.ToString(); }
                s += $" and [Дата Рождения] = '{yyyy}-{mm}-{dd}'"; MessageBox.Show(s); }
            if(TB14.Text.Length != 0) { s += $" and [Контактные данные] = '{TB14.Text}'"; }
            if(TB15.Text.Length != 0) { s += $" and [Паспортные данные] = '{TB15.Text}'"; }
            dt_user = Select(s);
            DG1.ItemsSource = dt_user.DefaultView;
            (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd:MM:yyyy";
            DG1.Columns[0].Visibility = Visibility.Hidden;
        }

        private void B_PrintWord_Br_Click(object sender, RoutedEventArgs e)
        {
            if (id_br_red != -1)
            {
                DataRowView row = (DataRowView)DG1.Items[id_br_red];
                dt_user = Select($"SELECT Номер.Номер,Номер.Класс FROM Номер,Бронь WHERE Номер.ID_Номера = Бронь.ID_Номер and Бронь.ID_Брони = {id_br}");
                string sourcePath = @"C:\Users\USER\Desktop\Практика ТРПО\Maket.docx";
                string path = @"C:\Users\USER\Desktop\Практика ТРПО\" + row["ID_Брони"] + ".docx";

                File.Copy(sourcePath, path, true);

                Word.Application wordApp = new Word.Application();
                wordApp.Visible = false;

                Word.Document wordDoucment = wordApp.Documents.Open(path);
                Word.Range range = wordDoucment.Content;
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{Client}", ReplaceWith: (string)row["Клиент"]);
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{Administrator}", ReplaceWith: (string)row["Администратор"]);
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{date_start}", ReplaceWith: row["Дата начала аренды"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{date_end}", ReplaceWith: row["Дата окончания аренды"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{Nomer}", ReplaceWith: dt_user.Rows[0]["Номер"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{class}", ReplaceWith: dt_user.Rows[0]["Класс"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{rent_cost}", ReplaceWith: row["Стоимость аренды"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{full_cost}", ReplaceWith: row["Стоимость аренды"].ToString());
                range = wordDoucment.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "{Usluga}", ReplaceWith: LabelCostAllUs.Content);

                wordDoucment.Save();
                wordApp.Quit();
                MessageBox.Show("Чек создан");
            }
        }

        private void B_PrintExcel_Br_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range rangeToHoldHyperlink;
            Excel.Range CellInstance;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.DisplayAlerts = false;
            rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
            CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
            int j = 1;
                foreach (DataRowView row in DG1.Items)
                {

                    for (int i = 1; i < DG1.Columns.Count; i++)
                    {
                        Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                        Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                        Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                        Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                        Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                        Excel.Range Range7 = xlWorkSheet.get_Range("F1");

                        Range2.ColumnWidth = 35;
                        Range2.EntireRow.AutoFit();
                        Range3.ColumnWidth = 35;
                        Range3.EntireRow.AutoFit();
                        Range4.ColumnWidth = 20;
                        Range4.EntireRow.AutoFit();
                        Range5.ColumnWidth = 20;
                        Range5.EntireRow.AutoFit();
                        Range6.ColumnWidth = 22;
                        Range6.EntireRow.AutoFit();
                        Range7.ColumnWidth = 20;
                        Range7.EntireRow.AutoFit();
                        
                        xlWorkSheet.Cells[1, 1] = "Администратор";
                        xlWorkSheet.Cells[1, 2] = "Клиент";
                        xlWorkSheet.Cells[1, 3] = "Номер";
                        xlWorkSheet.Cells[1, 4] = "Дата начала аренды";
                        xlWorkSheet.Cells[1, 5] = "Дата окончания аренды";
                        xlWorkSheet.Cells[1, 6] = "Стоимость";
                        
                        string[] conv;
                        xlWorkSheet.Cells[j + 1, i] = row[i].ToString();

                }
                    j++;
                }
            Excel.Range Range1 = xlWorkSheet.get_Range("A1");
            Range1.EntireRow.Font.Size = 14;
            Range1.EntireRow.AutoFit();
            Excel.Range tRange = xlWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            xlWorkBook.SaveAs(@"C:\Users\USER\Desktop\Практика ТРПО\Info.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            MessageBox.Show("Операция завершена");
        
        }

        private void B_Del_Client_Click(object sender, RoutedEventArgs e)
        {
            if (id_client_red != -1)
            {
                MessageBoxResult mb = MessageBox.Show("Вы действительно хотите удалить запись?", "Удаление записи", MessageBoxButton.YesNo);
                if (mb == MessageBoxResult.Yes)
                {
                    SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "UPDATE Клиент SET Статус = @1 WHERE ID_Клиент = @2";
                    cmd.Parameters.AddWithValue("@1", "Delete");
                    cmd.Parameters.AddWithValue("@2", id_client);
                    cmd.Connection = sqlConnection;
                    sqlConnection.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection.Close();

                    dt_user1 = Select("SELECT ID_Клиент,Фамилия,Имя,Отчество,[Контактные данные],[Паспортные данные],Адрес FROM [dbo].[Клиент] WHERE Статус is Null");
                    DG1.ItemsSource = dt_user1.DefaultView;
                    DG1.Columns[0].Visibility = Visibility.Hidden;
                }
            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }

        private void B_Del_Us_Click(object sender, RoutedEventArgs e)
        {
            if (id_usl_red != -1)
            {
                MessageBoxResult mb = MessageBox.Show("Вы действительно хотите удалить запись?", "Удаление записи", MessageBoxButton.YesNo);
                if (mb == MessageBoxResult.Yes)
                {
                    SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "UPDATE Услуга SET Статус = @1 WHERE ID_Услуги = @2";
                    cmd.Parameters.AddWithValue("@1", "Delete");
                    cmd.Parameters.AddWithValue("@2", id_uslug);
                    cmd.Connection = sqlConnection;
                    sqlConnection.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection.Close();

                    dt_user1 = Select("SELECT ID_Услуги,Наименование, Стоимость, [Время предоставления] FROM [dbo].[Услуга] WHERE Статус is Null");
                    DG1.ItemsSource = dt_user1.DefaultView;
                    DG1.Columns[0].Visibility = Visibility.Hidden;
                    (DG1.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
                }
            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }
        int id_spisuslug = -1;
        private void B_Del_UsinBr_Click(object sender, RoutedEventArgs e)
        {
            if (id_spisuslug != -1)
            {
                MessageBoxResult mb = MessageBox.Show("Вы действительно хотите удалить запись?", "Удаление записи", MessageBoxButton.YesNo);
                if (mb == MessageBoxResult.Yes)
                {
                    SqlConnection sqlConnection = new SqlConnection("server=VLADE;Trusted_Connection=Yes;DataBase=TRPO;");
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "UPDATE [Список оказанных услуг] SET Статус = @1 WHERE ID_Списка = @2";
                    cmd.Parameters.AddWithValue("@1", "Delete");
                    cmd.Parameters.AddWithValue("@2", id_spisuslug);
                    cmd.Connection = sqlConnection;
                    sqlConnection.Open();
                    cmd.ExecuteNonQuery();
                    sqlConnection.Close();

                    dt_user = Select($"SELECT [Список оказанных услуг].ID_Списка,Услуга.Наименование,Услуга.Стоимость,[Список оказанных услуг].[Количество услуг] FROM [Список оказанных услуг],Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].ID_Брони = {id_br} and [Список оказанных услуг].Статус is Null");
                    DG2.ItemsSource = dt_user.DefaultView;
                    (DG2.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
                    DG2.Columns[0].Visibility = Visibility.Hidden;
                    FullCost();
                    UsCost = Select("SELECT [Список оказанных услуг].[ID_Брони],(Sum(Услуга.Стоимость * [Список оказанных услуг].[Количество услуг])) as Стоимость FROM[Список оказанных услуг], Услуга WHERE Услуга.ID_Услуги = [Список оказанных услуг].ID_Услуги and [Список оказанных услуг].Статус is Null GROUP BY[Список оказанных услуг].ID_Брони");
                    foreach (DataRow dr in UsCost.Rows)
                    {
                        if ((int)dr["ID_Брони"] == id_br)
                        {
                            LabelCostAllUs.Content = dr["Стоимость"];
                            LabelCostAllUs.ContentStringFormat = "f";
                            break;
                        }
                        LabelCostAllUs.Content = "";
                    }
                } 
            }
            else
            {
                MessageBox.Show("Некорректно выделена строка");
            }
        }

        private void B_Filter_Client_Click(object sender, RoutedEventArgs e)
        {
            string s = "SELECT ID_Клиент, Фамилия, Имя, Отчество,[Контактные данные],[Паспортные данные],Адрес FROM Клиент WHERE Статус is NULL";
            if (TB1.Text.Length != 0) { s += $" and Фамилия = '{TB1.Text}'"; }
            if (TB2.Text.Length != 0) { s += $" and Имя = '{TB2.Text}'"; }
            if (TB3.Text.Length != 0) { s += $" and Отчество = '{TB3.Text}'"; }
            if (TB4.Text.Length != 0) { s += $" and [Контактные данные] = '{TB4.Text}'"; }
            if (TB5.Text.Length != 0) { s += $" and [Контактные данные] = '{TB5.Text}'"; }
            if (TB6.Text.Length != 0) { s += $" and [Паспортные данные] = '{TB6.Text}'"; }
            dt_user = Select(s);
            DG1.ItemsSource = dt_user.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
        }

        private void B_Filter_Nomer_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult mb = MessageBox.Show("Учитывать выбранный класс при фильтрации?","Фильтрация",MessageBoxButton.YesNo);
            string s = "SELECT ID_Номера,Номер,Этаж,[Количество комнат],Класс,Стоимость,Примечание FROM Номер WHERE Статус is NULL";
                if (TB16.Text.Length != 0 && TB16.Text != "0") { s += $" and Этаж = '{TB16.Text}'"; }
                if (TB17.Text.Length != 0 && TB17.Text != "0") { s += $" and [Количество комнат] = '{TB17.Text}'"; }
            if (mb == MessageBoxResult.Yes)
            {
                s += $" and Класс = '{CB1.Text}'";
                dt_user = Select(s);
                DG1.ItemsSource = dt_user.DefaultView;
            }
            if(mb == MessageBoxResult.No)
            {
                dt_user = Select(s);
                DG1.ItemsSource = dt_user.DefaultView;
            }
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "f";
        }

        private void B_Filter_Usl_Click(object sender, RoutedEventArgs e)
        {
            string s = "SELECT ID_Услуги,Наименование,Стоимость,[Время предоставления] FROM Услуга WHERE Статус is NULL";
            if (TB7.Text.Length != 0) { s += $" and Наименование = '{TB7.Text}'"; }
            if (TB8.Text.Length != 0) { s += $" and Стоимость = '{TB8.Text}'"; }
            if (TB9.Text.Length != 0) { s += $" and [Время предоставления] = '{TB9.Text}'"; }
            dt_user = Select(s);
            DG1.ItemsSource = dt_user.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[2] as DataGridTextColumn).Binding.StringFormat = "f";
        }

        private void B_Filter_Br_Click(object sender, RoutedEventArgs e)
        {
            string s = "SELECT Бронь.ID_Брони, (Администратор.Фамилия+' '+Администратор.Имя+' '+Администратор.Отчество) as Администратор,(Клиент.Фамилия+' '+Клиент.Имя+' '+Клиент.Отчество) as Клиент,Номер.Номер,Бронь.[Дата начала аренды],Бронь.[Дата окончания аренды],Бронь.[Стоимость аренды] FROM [dbo].[Бронь], [dbo].[Администратор],[dbo].[Клиент],[dbo].[Номер] WHERE Администратор.ID_Администратора = Бронь.ID_Администратора and Клиент.ID_Клиент = Бронь.ID_Клиент and Номер.ID_Номера = Бронь.ID_Номер and Бронь.Статус is Null";
            if (TB18.Text.Length != 0) { s += $" and Клиент.Фамилия Like '%{TB18.Text}%'"; }
            if (TB22.Text.Length != 0) { s += $" and Клиент.Имя Like '%{TB22.Text}%'"; }
            if (TB23.Text.Length != 0) { s += $" and Клиент.Отчество Like '%{TB23.Text}%'"; }
            if (DP1.Text.Length != 0) { s += $" and Бронь.[Дата начала аренды] = '{DP1.Text}'"; }
            if (DP2.Text.Length != 0) { s += $" and Бронь.[Дата окончания аренды] = '{DP2.Text}'"; }
            dt_user = Select(s);
            DG1.ItemsSource = dt_user.DefaultView;
            DG1.Columns[0].Visibility = Visibility.Hidden;
            (DG1.Columns[4] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
            (DG1.Columns[5] as DataGridTextColumn).Binding.StringFormat = "dd.MM.yyyy hh:mm";
            (DG1.Columns[6] as DataGridTextColumn).Binding.StringFormat = "f";
            FullCost();
        }

        private void B_PrintExcel_Client_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range rangeToHoldHyperlink;
            Excel.Range CellInstance;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.DisplayAlerts = false;
            rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
            CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
            int j = 1;
            foreach (DataRowView row in DG1.Items)
            {

                for (int i = 1; i < DG1.Columns.Count; i++)
                {
                    Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                    Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                    Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                    Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                    Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                    Excel.Range Range7 = xlWorkSheet.get_Range("F1");

                    Range2.ColumnWidth = 35;
                    Range2.EntireRow.AutoFit();
                    Range3.ColumnWidth = 35;
                    Range3.EntireRow.AutoFit();
                    Range4.ColumnWidth = 20;
                    Range4.EntireRow.AutoFit();
                    Range5.ColumnWidth = 20;
                    Range5.EntireRow.AutoFit();
                    Range6.ColumnWidth = 22;
                    Range6.EntireRow.AutoFit();
                    Range7.ColumnWidth = 20;
                    Range7.EntireRow.AutoFit();

                    xlWorkSheet.Cells[1, 1] = "Фамилия";
                    xlWorkSheet.Cells[1, 2] = "Имя";
                    xlWorkSheet.Cells[1, 3] = "Отчество";
                    xlWorkSheet.Cells[1, 4] = "Контактные данные";
                    xlWorkSheet.Cells[1, 5] = "Паспортные данные";
                    xlWorkSheet.Cells[1, 6] = "Адрес";

                    string[] conv;
                    xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                    (xlWorkSheet.Cells[j+1, 4] as Excel.Range).NumberFormat = "@";
                }
                j++;
            }
            Excel.Range Range1 = xlWorkSheet.get_Range("A1");
            Range1.EntireRow.Font.Size = 14;
            Range1.EntireRow.AutoFit();
            Excel.Range tRange = xlWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            xlWorkBook.SaveAs(@"C:\Users\USER\Desktop\Практика ТРПО\InfoClient.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            MessageBox.Show("Операция завершена");
        }

        private void B_PrintExcel_Adm_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range rangeToHoldHyperlink;
            Excel.Range CellInstance;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.DisplayAlerts = false;
            rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
            CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
            int j = 1;
            foreach (DataRowView row in DG1.Items)
            {

                for (int i = 1; i < DG1.Columns.Count; i++)
                {
                    Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                    Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                    Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                    Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                    Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                    Excel.Range Range7 = xlWorkSheet.get_Range("F1");

                    Range2.ColumnWidth = 35;
                    Range2.EntireRow.AutoFit();
                    Range3.ColumnWidth = 35;
                    Range3.EntireRow.AutoFit();
                    Range4.ColumnWidth = 20;
                    Range4.EntireRow.AutoFit();
                    Range5.ColumnWidth = 20;
                    Range5.EntireRow.AutoFit();
                    Range6.ColumnWidth = 22;
                    Range6.EntireRow.AutoFit();
                    Range7.ColumnWidth = 20;
                    Range7.EntireRow.AutoFit();

                    xlWorkSheet.Cells[1, 1] = "Фамилия";
                    xlWorkSheet.Cells[1, 2] = "Имя";
                    xlWorkSheet.Cells[1, 3] = "Отчество";
                    xlWorkSheet.Cells[1, 4] = "Дата рождения";
                    xlWorkSheet.Cells[1, 5] = "Контактные данные";
                    xlWorkSheet.Cells[1, 6] = "Паспортные данные";

                    string[] conv;
                    xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                    (xlWorkSheet.Cells[j + 1, 5] as Excel.Range).NumberFormat = "@";
                }
                j++;
            }
            Excel.Range Range1 = xlWorkSheet.get_Range("A1");
            Range1.EntireRow.Font.Size = 14;
            Range1.EntireRow.AutoFit();
            Excel.Range tRange = xlWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            xlWorkBook.SaveAs(@"C:\Users\USER\Desktop\Практика ТРПО\InfoAdm.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            MessageBox.Show("Операция завершена");
        }

        private void B_PrintExcel_Usl_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range rangeToHoldHyperlink;
            Excel.Range CellInstance;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.DisplayAlerts = false;
            rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
            CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
            int j = 1;
            foreach (DataRowView row in DG1.Items)
            {

                for (int i = 1; i < DG1.Columns.Count; i++)
                {
                    Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                    Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                    Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                    Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                    Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                    Excel.Range Range7 = xlWorkSheet.get_Range("F1");

                    Range2.ColumnWidth = 35;
                    Range2.EntireRow.AutoFit();
                    Range3.ColumnWidth = 35;
                    Range3.EntireRow.AutoFit();
                    Range4.ColumnWidth = 30;
                    Range4.EntireRow.AutoFit();
                    Range5.ColumnWidth = 20;
                    Range5.EntireRow.AutoFit();
                    Range6.ColumnWidth = 22;
                    Range6.EntireRow.AutoFit();
                    Range7.ColumnWidth = 20;
                    Range7.EntireRow.AutoFit();

                    xlWorkSheet.Cells[1, 1] = "Наименование";
                    xlWorkSheet.Cells[1, 2] = "Стоимость";
                    xlWorkSheet.Cells[1, 3] = "Время предоставления";
                    

                    string[] conv;
                    xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                    (xlWorkSheet.Cells[j + 1, 2] as Excel.Range).NumberFormat = "@";
                }
                j++;
            }
            Excel.Range Range1 = xlWorkSheet.get_Range("A1");
            Range1.EntireRow.Font.Size = 14;
            Range1.EntireRow.AutoFit();
            Excel.Range tRange = xlWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            xlWorkBook.SaveAs(@"C:\Users\USER\Desktop\Практика ТРПО\InfoUsl.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            MessageBox.Show("Операция завершена");
        }

        private void B_PrintExcel_Nomer_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range rangeToHoldHyperlink;
            Excel.Range CellInstance;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.DisplayAlerts = false;
            rangeToHoldHyperlink = xlWorkSheet.get_Range("A1", Type.Missing);
            CellInstance = xlWorkSheet.get_Range("A1", Type.Missing);
            int j = 1;
            foreach (DataRowView row in DG1.Items)
            {

                for (int i = 1; i < DG1.Columns.Count; i++)
                {
                    Excel.Range Range2 = xlWorkSheet.get_Range("A1");
                    Excel.Range Range3 = xlWorkSheet.get_Range("B1");
                    Excel.Range Range4 = xlWorkSheet.get_Range("C1");
                    Excel.Range Range5 = xlWorkSheet.get_Range("D1");
                    Excel.Range Range6 = xlWorkSheet.get_Range("E1");
                    Excel.Range Range7 = xlWorkSheet.get_Range("F1");

                    Range2.ColumnWidth = 35;
                    Range2.EntireRow.AutoFit();
                    Range3.ColumnWidth = 35;
                    Range3.EntireRow.AutoFit();
                    Range4.ColumnWidth = 30;
                    Range4.EntireRow.AutoFit();
                    Range5.ColumnWidth = 20;
                    Range5.EntireRow.AutoFit();
                    Range6.ColumnWidth = 22;
                    Range6.EntireRow.AutoFit();
                    Range7.ColumnWidth = 20;
                    Range7.EntireRow.AutoFit();

                    xlWorkSheet.Cells[1, 1] = "Номер";
                    xlWorkSheet.Cells[1, 2] = "Этаж";
                    xlWorkSheet.Cells[1, 3] = "Кол-во комнат";
                    xlWorkSheet.Cells[1, 4] = "Класс";
                    xlWorkSheet.Cells[1, 5] = "Стоимость";
                    xlWorkSheet.Cells[1, 6] = "Примечание";
                    xlWorkSheet.Cells[1, 5] = "Стоимость";

                    string[] conv;
                    xlWorkSheet.Cells[j + 1, i] = row[i].ToString();
                    (xlWorkSheet.Cells[j + 1, 5] as Excel.Range).NumberFormat = "@";
                }
                j++;
            }
            Excel.Range Range1 = xlWorkSheet.get_Range("A1");
            Range1.EntireRow.Font.Size = 14;
            Range1.EntireRow.AutoFit();
            Excel.Range tRange = xlWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            xlWorkBook.SaveAs(@"C:\Users\USER\Desktop\Практика ТРПО\InfoNom.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            MessageBox.Show("Операция завершена");
        }
    }
}
