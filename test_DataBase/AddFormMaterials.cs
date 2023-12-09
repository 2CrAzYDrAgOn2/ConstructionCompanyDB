using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormMaterials : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormMaterials()
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
                var materialName = textBoxMaterialName.Text;
                if (int.TryParse(textBoxUnitPrice.Text, out int unitPrice) && int.TryParse(textBoxQuantinityInStock.Text, out int quantinityInStock))
                {
                    var addQuery = $"insert into Materials (MaterialName, UnitPrice, QuantityInStock) values ('{materialName}', '{unitPrice}', '{quantinityInStock}')";
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