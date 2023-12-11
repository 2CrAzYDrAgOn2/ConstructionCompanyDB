using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class SignUp : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public SignUp()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        /// <summary>
        /// SignUp_Load вызывается при загрузке формы "SignUp"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SignUp_Load(object sender, EventArgs e)
        {
            textBoxPassword.PasswordChar = '•';
            textBoxLogin.MaxLength = 50;
            textBoxPassword.MaxLength = 50;
        }

        /// <summary>
        /// ButtonCreate_Click вызывается при нажатии на кнопку "Войти"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonCreate_Click(object sender, EventArgs e)
        {
            var login = textBoxLogin.Text;
            var password = textBoxPassword.Text;
            string querystring = $"Insert into Registration(UserLogin, UserPassword, IsAdmin) values('{login}', '{password}', 0)";
            SqlCommand sqlCommand = new SqlCommand(querystring, dataBase.GetConnection());
            dataBase.OpenConnection();
            if (sqlCommand.ExecuteNonQuery() == 1)
            {
                MessageBox.Show("Аккаунт успешно создан!", "Успех!");
                LogIn formLogin = new LogIn();
                this.Hide();
                formLogin.ShowDialog();
            }
            else
            {
                MessageBox.Show("Аккаунт не создан!");
            }
            dataBase.CloseConnection();
        }

        /// <summary>
        /// ButtonClear_Click вызывается при нажатии на кнопку очистки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClear_Click(object sender, EventArgs e)
        {
            textBoxLogin.Text = "";
            textBoxPassword.Text = "";
        }
    }
}