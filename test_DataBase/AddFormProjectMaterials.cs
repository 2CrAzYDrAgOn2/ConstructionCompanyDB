using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormProjectMaterials : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormProjectMaterials()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void ButtonSave_Click(object sender, EventArgs e)
        {
            try
            {
                dataBase.OpenConnection();
                if (int.TryParse(textBoxProjectIDProjectMaterials.Text, out int projectIDProjectMaterials) && int.TryParse(textBoxMaterialIDProjectMaterials.Text, out int materialIDProjectMaterials) && int.TryParse(textBoxQuantinityUsed.Text, out int quantinityUsed))
                {
                    var addQuery = $"insert into ProjectMaterials (ProjectID, MaterialID, QuantityUsed) values ('{projectIDProjectMaterials}', '{materialIDProjectMaterials}', '{quantinityUsed}')";
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