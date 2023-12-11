using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormEmployees : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormEmployees()
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
                var firstName = textBoxFirstName.Text;
                var lastName = textBoxLastName.Text;
                var position = textBoxPosition.Text;
                var hireDate = textBoxHireDate.Value;
                var employees = textBoxEmailEmployees.Text;
                var phoneNumber = textBoxPhoneNumber.Text;
                if (int.TryParse(textBoxSalary.Text, out int salary))
                {
                    var addQuery = $"insert into Employees (FirstName, LastName, Position, HireDate, Salary, Email, PhoneNumber) values ('{firstName}', '{lastName}', '{position}', '{hireDate}', '{salary}', '{employees}', '{phoneNumber}')";
                    var sqlCommand = new SqlCommand(addQuery, dataBase.GetConnection());
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно создана!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Цена должна иметь числовой формат!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
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