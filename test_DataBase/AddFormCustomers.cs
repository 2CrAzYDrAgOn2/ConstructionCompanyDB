using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace test_DataBase
{
    public partial class AddFormCustomers : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormCustomers()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        /// <summary>
        /// ButtonSave_Click вызывается при нажатии на кнопку "Сохранить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSave_Click(object sender, EventArgs e)
        {
            try
            {
                dataBase.OpenConnection();
                var customerName = textBoxCustomerName.Text;
                var contactPerson = textBoxContactPerson.Text;
                var contactNumber = textBoxContactNumber.Text;
                var email = textBoxEmail.Text;
                var addQuery = $"insert into Customers (CustomerName, ContactPerson, ContactNumber, Email) values ('{customerName}', '{contactPerson}', '{contactNumber}', '{email}')";
                var sqlCommand = new SqlCommand(addQuery, dataBase.GetConnection());
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Запись успешно создана!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dataBase.CloseConnection();
            }
        }
    }
}