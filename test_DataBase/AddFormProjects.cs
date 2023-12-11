using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormProjects : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormProjects()
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
                var projectName = textBoxProjectName.Text;
                var startDate = textBoxStartDate.Value;
                var endDate = textBoxEndDate.Value;
                var status = textBoxStatus.Text;
                if (int.TryParse(textBoxBudget.Text, out int budget))
                {
                    var addQuery = $"insert into Projects (ProjectName, StartDate, EndDate, Budget, Status) values ('{projectName}', '{startDate}', '{endDate}', '{budget}', '{status}')";
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