using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
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
using System.Windows.Shapes;


namespace Практика_ТРПО
{
    /// <summary>
    /// Логика взаимодействия для Authorization.xaml
    /// </summary>
    public partial class Authorization : Window
    {
        public Authorization()
        {
            InitializeComponent();
            
            dt_user = Select("SELECT ID_Администратора,Фамилия,Имя,Отчество,Login,Password FROM Администратор WHERE Login is not NULL or Password is not NULL");
        }
        DataTable dt_user;
        private void B1_Click(object sender, RoutedEventArgs e)
        {
            bool ok = false;
            foreach(DataRow a in dt_user.Rows)
            {
                if(Login.Text == a["Login"].ToString())
                {
                    if(Password.Text == a["Password"].ToString())
                    {
                        
                        int id_admina_from_reg = (int)a[0];
                        string admin = a[1].ToString() + " " + (a[2].ToString())[0] + "." + (a[3].ToString())[0]+".";
                        MainWindow mw = new MainWindow(id_admina_from_reg,admin);
                        mw.ShowDialog();
                        Hide();
                        ok = true;
                        break;
                    }
                }
            }
            if (ok == false)
            {
                MessageBox.Show("Ошибка авторизации. Проверьте правильность введённых данных");
                Login.Text = "";
                Password.Text = "";
            }
            
        }

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
        
    }
}
