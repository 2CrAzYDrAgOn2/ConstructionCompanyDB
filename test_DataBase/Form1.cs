using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout.Properties;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_DataBase
{
    internal enum RowState
    {
        Existed,
        New,
        Modified,
        ModifiedNew,
        Deleted
    }

    public partial class Form1 : Form
    {
        private readonly DataBase dataBase = new DataBase();
        private bool admin;
        private int selectedRow;

        public Form1()
        {
            try
            {
                InitializeComponent();
                StartPosition = FormStartPosition.CenterScreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// SetAdminStatus проверяет доступ
        /// </summary>
        /// <param name="isAdmin"></param>
        public void SetAdminStatus(bool isAdmin)
        {
            admin = isAdmin;
        }

        /// <summary>
        /// CreateColumns вызывается при создании колонок
        /// </summary>
        private void CreateColumns()
        {
            try
            {
                dataGridViewProjects.Columns.Add("ProjectID", "Номер");
                dataGridViewProjects.Columns.Add("ProjectName", "Название проекта");
                dataGridViewProjects.Columns.Add("StartDate", "Дата начала");
                dataGridViewProjects.Columns.Add("EndDate", "Дата конца");
                dataGridViewProjects.Columns.Add("Budget", "Бюджет");
                dataGridViewProjects.Columns.Add("Status", "Статус");
                dataGridViewProjects.Columns.Add("IsNew", String.Empty);
                dataGridViewCustomers.Columns.Add("CustomerID", "Номер");
                dataGridViewCustomers.Columns.Add("CustomerName", "Имя клиента");
                dataGridViewCustomers.Columns.Add("ContactPerson", "Контактное лицо");
                dataGridViewCustomers.Columns.Add("ContactNumber", "Контактный номер");
                dataGridViewCustomers.Columns.Add("Email", "Email");
                dataGridViewCustomers.Columns.Add("IsNew", String.Empty);
                dataGridViewEmployees.Columns.Add("EmployeeID", "Номер");
                dataGridViewEmployees.Columns.Add("FirstName", "Имя");
                dataGridViewEmployees.Columns.Add("LastName", "Фамилия");
                dataGridViewEmployees.Columns.Add("Position", "Должность");
                dataGridViewEmployees.Columns.Add("HireDate", "Дата найма");
                dataGridViewEmployees.Columns.Add("Salary", "Зарплата");
                dataGridViewEmployees.Columns.Add("Email", "Email");
                dataGridViewEmployees.Columns.Add("PhoneNumber", "Номер телефона");
                dataGridViewEmployees.Columns.Add("IsNew", String.Empty);
                dataGridViewMaterials.Columns.Add("MaterialID", "Номер");
                dataGridViewMaterials.Columns.Add("MaterialName", "Название материала");
                dataGridViewMaterials.Columns.Add("UnitPrice", "Цена");
                dataGridViewMaterials.Columns.Add("QuantinityInStock", "В наличии");
                dataGridViewMaterials.Columns.Add("IsNew", String.Empty);
                dataGridViewProjectMaterials.Columns.Add("ProjectID", "Номер проекта");
                dataGridViewProjectMaterials.Columns.Add("MaterialID", "Номер материала");
                dataGridViewProjectMaterials.Columns.Add("QuantinityUsed", "Использовано");
                dataGridViewProjectMaterials.Columns.Add("IsNew", String.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// CreateColumns вызывается при очистке полей
        /// </summary>
        private void ClearFields()
        {
            try
            {
                textBoxProjectID.Text = "";
                textBoxProjectName.Text = "";
                textBoxStartDate.Text = "";
                textBoxEndDate.Text = "";
                textBoxBudget.Text = "";
                textBoxStatus.Text = "";
                textBoxCustomerID.Text = "";
                textBoxCustomerName.Text = "";
                textBoxContactPerson.Text = "";
                textBoxContactNumber.Text = "";
                textBoxEmail.Text = "";
                textBoxEmployeeID.Text = "";
                textBoxFirstName.Text = "";
                textBoxLastName.Text = "";
                textBoxPosition.Text = "";
                textBoxHireDate.Text = "";
                textBoxSalary.Text = "";
                textBoxEmailEmployees.Text = "";
                textBoxPhoneNumber.Text = "";
                textBoxMaterialID.Text = "";
                textBoxMaterialName.Text = "";
                textBoxUnitPrice.Text = "";
                textBoxQuantinityInStock.Text = "";
                textBoxProjectIDProjectMaterials.Text = "";
                textBoxMaterialIDProjectMaterials.Text = "";
                textBoxQuantinityUsed.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ReadSingleRow вызывается при чтении строк
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="iDataRecord"></param>
        private void ReadSingleRow(DataGridView dataGridView, IDataRecord iDataRecord)
        {
            try
            {
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetDateTime(2), iDataRecord.GetDateTime(3), iDataRecord.GetInt32(4), iDataRecord.GetString(5), RowState.Modified);
                        break;

                    case "dataGridViewCustomers":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewEmployees":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetDateTime(4), iDataRecord.GetInt32(5), iDataRecord.GetString(6), iDataRecord.GetString(7), RowState.Modified);
                        break;

                    case "dataGridViewMaterials":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetInt32(2), iDataRecord.GetInt32(3), RowState.Modified);
                        break;

                    case "dataGridViewProjectMaterials":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), iDataRecord.GetInt32(2), RowState.Modified);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// RefreshDataGrid вызывается при обновлении dataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="tableName"></param>
        private void RefreshDataGrid(DataGridView dataGridView, string tableName)
        {
            try
            {
                dataGridView.Rows.Clear();
                string queryString = $"select * from {tableName}";
                SqlCommand sqlCommand = new SqlCommand(queryString, dataBase.GetConnection());
                dataBase.OpenConnection();
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                while (sqlDataReader.Read())
                {
                    ReadSingleRow(dataGridView, sqlDataReader);
                }
                sqlDataReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Form1_Load вызывается при загрузке формы "Form1"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                CreateColumns();
                RefreshDataGrid(dataGridViewProjects, "Projects");
                RefreshDataGrid(dataGridViewCustomers, "Customers");
                RefreshDataGrid(dataGridViewEmployees, "Employees");
                RefreshDataGrid(dataGridViewMaterials, "Materials");
                RefreshDataGrid(dataGridViewProjectMaterials, "ProjectMaterials");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// DataGridView_CellClick вызывается при нажатии на ячейку в DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="selectedRow"></param>
        private void DataGridView_CellClick(DataGridView dataGridView, int selectedRow)
        {
            try
            {
                DataGridViewRow dataGridViewRow = dataGridView.Rows[selectedRow];
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        textBoxProjectID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxProjectName.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxStartDate.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxEndDate.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxBudget.Text = dataGridViewRow.Cells[4].Value.ToString();
                        textBoxStatus.Text = dataGridViewRow.Cells[5].Value.ToString();
                        break;

                    case "dataGridViewCustomers":
                        textBoxCustomerID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxCustomerName.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxContactPerson.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxContactNumber.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxEmail.Text = dataGridViewRow.Cells[4].Value.ToString();
                        break;

                    case "dataGridViewEmployees":
                        textBoxEmployeeID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxFirstName.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxLastName.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxPosition.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxHireDate.Text = dataGridViewRow.Cells[4].Value.ToString();
                        textBoxSalary.Text = dataGridViewRow.Cells[5].Value.ToString();
                        textBoxEmailEmployees.Text = dataGridViewRow.Cells[6].Value.ToString();
                        textBoxPhoneNumber.Text = dataGridViewRow.Cells[7].Value.ToString();
                        break;

                    case "dataGridViewMaterials":
                        textBoxMaterialID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxMaterialName.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxUnitPrice.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxQuantinityInStock.Text = dataGridViewRow.Cells[3].Value.ToString();
                        break;

                    case "dataGridViewProjectMaterials":
                        textBoxProjectIDProjectMaterials.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxMaterialIDProjectMaterials.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxQuantinityUsed.Text = dataGridViewRow.Cells[2].Value.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Search вызывается при поиске данных в DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        private void Search(DataGridView dataGridView)
        {
            try
            {
                dataGridView.Rows.Clear();
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        string searchStringProjects = $"select * from Projects where concat (ProjectID, ProjectName, StartDate, EndDate, Budget, Status) like '%" + textBoxSearchProjects.Text + "%'";
                        SqlCommand sqlCommandProjects = new SqlCommand(searchStringProjects, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderProjects = sqlCommandProjects.ExecuteReader();
                        while (sqlDataReaderProjects.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderProjects);
                        }
                        sqlDataReaderProjects.Close();
                        break;

                    case "dataGridViewCustomers":
                        string searchStringCustomers = $"select * from Customers where concat (CustomerID, CustomerName, ContactPerson, ContactNumber, Email) like '%" + textBoxSearchCustomers.Text + "%'";
                        SqlCommand sqlCommandCustomers = new SqlCommand(searchStringCustomers, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderCustomers = sqlCommandCustomers.ExecuteReader();
                        while (sqlDataReaderCustomers.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderCustomers);
                        }
                        sqlDataReaderCustomers.Close();
                        break;

                    case "dataGridViewEmployees":
                        string searchStringEmployees = $"select * from Employees where concat (EmployeeID, FirstName, LastName, Position, HireDate, Salary, Email, PhoneNumber) like '%" + textBoxSearchEmployees.Text + "%'";
                        SqlCommand sqlCommandEmployees = new SqlCommand(searchStringEmployees, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderEmployees = sqlCommandEmployees.ExecuteReader();
                        while (sqlDataReaderEmployees.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderEmployees);
                        }
                        sqlDataReaderEmployees.Close();
                        break;

                    case "dataGridViewMaterials":
                        string searchStringMaterials = $"select * from Materials where concat (MaterialID, MaterialName, UnitPrice, QuantinityInStock) like '%" + textBoxSearchMaterials.Text + "%'";
                        SqlCommand sqlCommandMaterials = new SqlCommand(searchStringMaterials, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderMaterials = sqlCommandMaterials.ExecuteReader();
                        while (sqlDataReaderMaterials.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderMaterials);
                        }
                        sqlDataReaderMaterials.Close();
                        break;

                    case "dataGridViewProjectMaterials":
                        string searchStringProjectMaterials = $"select * from ProjectMaterials where concat (ProjectID, MaterialID, QuantinityUsed) like '%" + textBoxSearchProjectMaterials.Text + "%'";
                        SqlCommand sqlCommandProjectMaterials = new SqlCommand(searchStringProjectMaterials, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderProjectMaterials = sqlCommandProjectMaterials.ExecuteReader();
                        while (sqlDataReaderProjectMaterials.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderProjectMaterials);
                        }
                        sqlDataReaderProjectMaterials.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// DeleteRow вызывается при удалении строки
        /// </summary>
        /// <param name="dataGridView"></param>
        private void DeleteRow(DataGridView dataGridView)
        {
            try
            {
                int index = dataGridView.CurrentCell.RowIndex;
                dataGridView.Rows[index].Visible = false;
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                        break;

                    case "dataGridViewCustomers":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                        break;

                    case "dataGridViewEmployees":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[8].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[8].Value = RowState.Deleted;
                        break;

                    case "dataGridViewMaterials":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[4].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[4].Value = RowState.Deleted;
                        break;

                    case "dataGridViewProjectMaterials":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[3].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[3].Value = RowState.Deleted;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// UpdateBase вызывается при обновлении базы данных
        /// </summary>
        /// <param name="dataGridView"></param>
        private void UpdateBase(DataGridView dataGridView)
        {
            try
            {
                dataBase.OpenConnection();
                for (int index = 0; index < dataGridView.Rows.Count; index++)
                {
                    switch (dataGridView.Name)
                    {
                        case "dataGridViewProjects":
                            var rowStateProjects = (RowState)dataGridView.Rows[index].Cells[6].Value;
                            if (rowStateProjects == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateProjects == RowState.Deleted)
                            {
                                var projectID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Projects where ClientID = {projectID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateProjects == RowState.Modified)
                            {
                                var projectID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var projectName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var startDate = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var endDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var budget = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var status = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var changeQuery = $"update Projects set ProjectName = '{firstName}', StartDate = '{lastName}', EndDate = '{email}', Budget = '{phone}' where ProjectID = '{projectID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewCustomers":
                            var rowStateCustomers = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateCustomers == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateCustomers == RowState.Deleted)
                            {
                                var tourID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Customers where TourID = {tourID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateCustomers == RowState.Modified)
                            {
                                var tourID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var tourName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var destination = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var startDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var endDate = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update Customers set TourName = '{tourName}', Destination = '{destination}', StartDate = '{startDate}', EndDate = '{endDate}', Price = '{price}' where TourID = '{tourID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewEmployees":
                            var rowStateEmployees = (RowState)dataGridView.Rows[index].Cells[8].Value;
                            if (rowStateEmployees == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEmployees == RowState.Deleted)
                            {
                                var bookingID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Employees where BookingID = {bookingID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEmployees == RowState.Modified)
                            {
                                var bookingID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var clientID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var tourID = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var bookingDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var numberOfPersons = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var totalAmount = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var totalAmount = dataGridView.Rows[index].Cells[6].Value.ToString();
                                var totalAmount = dataGridView.Rows[index].Cells[7].Value.ToString();
                                var changeQuery = $"update Employees set ClientID = '{clientID}', TourID = '{tourID}', BookingDate = '{bookingDate}', NumberOfPersons = '{numberOfPersons}', TotalAmount = '{totalAmount}' where BookingID = '{bookingID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewMaterials":
                            var rowStateMaterials = (RowState)dataGridView.Rows[index].Cells[4].Value;
                            if (rowStateMaterials == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateMaterials == RowState.Deleted)
                            {
                                var paymentID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Materials where PaymentID = {paymentID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateMaterials == RowState.Modified)
                            {
                                var paymentID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var bookingID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var paymentDate = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var amount = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var changeQuery = $"update Materials set BookingID = '{bookingID}', PaymentDate = '{paymentDate}', Amount = '{amount}' where PaymentID = '{paymentID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewProjectMaterials":
                            var rowStateProjectMaterials = (RowState)dataGridView.Rows[index].Cells[3].Value;
                            if (rowStateProjectMaterials == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateProjectMaterials == RowState.Deleted)
                            {
                                var paymentID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from ProjectMaterials where PaymentID = {paymentID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateProjectMaterials == RowState.Modified)
                            {
                                var paymentID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var bookingID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var paymentDate = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var changeQuery = $"update ProjectMaterials set BookingID = '{bookingID}', PaymentDate = '{paymentDate}', Amount = '{amount}' where PaymentID = '{paymentID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;
                    }
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

        /// <summary>
        /// Change вызывается при изменении данных в базе данных
        /// </summary>
        /// <param name="dataGridView"></param>
        private void Change(DataGridView dataGridView)
        {
            try
            {
                var selectedRowIndex = dataGridView.CurrentCell.RowIndex;
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        var projectID = textBoxProjectID.Text;
                        var projectName = textBoxProjectName.Text;
                        var startDate = textBoxStartDate.Text;
                        var endDate = textBoxEndDate.Text;
                        var budget = textBoxBudget.Text;
                        var status = textBoxStatus.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(projectID, projectName, startDate, endDate, budget, status);
                        dataGridView.Rows[selectedRowIndex].Cells[6].Value = RowState.Modified;
                        break;

                    case "dataGridViewCustomers":
                        var customerID = textBoxCustomerID.Text;
                        var customerName = textBoxCustomerName.Text;
                        var contactPerson = textBoxContactPerson.Text;
                        var contactNumber = textBoxContactNumber.Text;
                        var email = textBoxEmail.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(customerID, customerName, contactPerson, contactNumber, email);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewEmployees":
                        var employeeID = textBoxEmployeeID.Text;
                        var firstName = textBoxFirstName.Text;
                        var lastName = textBoxLastName.Text;
                        var position = textBoxPosition.Text;
                        var hireDate = textBoxHireDate.Text;
                        var salary = textBoxSalary.Text;
                        var employees = textBoxEmailEmployees.Text;
                        var phoneNumber = textBoxPhoneNumber.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(employeeID, firstName, lastName, position, hireDate, salary, employees, phoneNumber);
                        dataGridView.Rows[selectedRowIndex].Cells[8].Value = RowState.Modified;
                        break;

                    case "dataGridViewMaterials":
                        var materialID = textBoxMaterialID.Text;
                        var materialName = textBoxMaterialName.Text;
                        var unitPrice = textBoxUnitPrice.Text;
                        var quantinityInStock = textBoxQuantinityInStock.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(materialID, materialName, unitPrice, quantinityInStock);
                        dataGridView.Rows[selectedRowIndex].Cells[4].Value = RowState.Modified;
                        break;

                    case "dataGridViewProjectMaterials":
                        var projectIDProjectMaterials = textBoxProjectIDProjectMaterials.Text;
                        var materialIDProjectMaterials = textBoxMaterialIDProjectMaterials.Text;
                        var quantinityUsed = textBoxQuantinityUsed.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(projectIDProjectMaterials, materialIDProjectMaterials, quantinityUsed);
                        dataGridView.Rows[selectedRowIndex].Cells[3].Value = RowState.Modified;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToWord вызывается при экспорте данных в Word
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToWord(DataGridView dataGridView)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                Paragraph title = doc.Paragraphs.Add();
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        title.Range.Text = "Данные проектов";
                        break;

                    case "dataGridViewCustomers":
                        title.Range.Text = "Данные клиентов";
                        break;

                    case "dataGridViewEmployees":
                        title.Range.Text = "Данные сотрудников";
                        break;

                    case "dataGridViewMaterials":
                        title.Range.Text = "Данные материалов";
                        break;

                    case "dataGridViewProjectMaterials":
                        title.Range.Text = "Данные использования материалов";
                        break;
                }
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();
                Table table = doc.Tables.Add(title.Range, dataGridView.RowCount + 1, dataGridView.ColumnCount - 1);
                for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                {
                    table.Cell(1, col + 1).Range.Text = dataGridView.Columns[col].HeaderText;
                }
                for (int row = 0; row < dataGridView.RowCount; row++)
                {
                    for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                    {
                        table.Cell(row + 2, col + 1).Range.Text = dataGridView[col, row].Value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToExcel вызывается при экспорте данных в Excel
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToExcel(DataGridView dataGridView)
        {
            try
            {
                var excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                string title = "";
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        title = "Данные проектов";
                        break;

                    case "dataGridViewCustomers":
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewEmployees":
                        title = "Данные сотрудников";
                        break;

                    case "dataGridViewMaterials":
                        title = "Данные материалов";
                        break;

                    case "dataGridViewProjectMaterials":
                        title = "Данные использования материалов";
                        break;
                }
                Excel.Range titleRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, dataGridView.ColumnCount - 1]];
                titleRange.Merge();
                titleRange.Value = title;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                for (int col = 0; col < dataGridView.ColumnCount; col++)
                {
                    worksheet.Cells[2, col + 1] = dataGridView.Columns[col].HeaderText;
                }
                for (int row = 0; row < dataGridView.RowCount; row++)
                {
                    for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                    {
                        worksheet.Cells[row + 3, col + 1] = dataGridView[col, row].Value.ToString();
                        Excel.Range dataRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[dataGridView.RowCount + 2, dataGridView.ColumnCount]];
                        dataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        dataRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                }
                worksheet.Columns.AutoFit();
                worksheet.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToPDF вызывается при экспорте данных в PDF
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToPDF(DataGridView dataGridView)
        {
            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.pdf");
                var pdfWriter = new PdfWriter(filePath);
                var pdfDocument = new PdfDocument(pdfWriter);
                var pdfDoc = new iText.Layout.Document(pdfDocument);
                PdfFont timesFont = PdfFontFactory.CreateFont("c:/windows/fonts/times.ttf", PdfEncodings.IDENTITY_H, true);
                string title = "";
                switch (dataGridView.Name)
                {
                    case "dataGridViewProjects":
                        title = "Данные проектов";
                        break;

                    case "dataGridViewCustomers":
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewEmployees":
                        title = "Данные сотрудников";
                        break;

                    case "dataGridViewMaterials":
                        title = "Данные материалов";
                        break;

                    case "dataGridViewProjectMaterials":
                        title = "Данные использования материалов";
                        break;
                }
                pdfDoc.Add(new iText.Layout.Element.Paragraph(title).SetFont(timesFont).SetTextAlignment(TextAlignment.CENTER));
                iText.Layout.Element.Table table = new iText.Layout.Element.Table(dataGridView.Columns.Count - 1);
                table.UseAllAvailableWidth();
                var columnsList = dataGridView.Columns.Cast<DataGridViewColumn>().ToList();
                foreach (DataGridViewColumn column in columnsList.Take(dataGridView.Columns.Count - 1))
                {
                    iText.Layout.Element.Cell headerCell = new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(column.HeaderText).SetFont(timesFont));
                    table.AddHeaderCell(headerCell);
                }
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells.Cast<DataGridViewCell>().Take(dataGridView.Columns.Count - 1))
                    {
                        table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(cell.Value.ToString()).SetFont(timesFont)));
                    }
                }
                pdfDoc.Add(table);
                pdfDoc.Close();
                MessageBox.Show("PDF успешно экспортирован.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ButtonRefresh_Click вызывается при нажатии на кнопку обновления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                RefreshDataGrid(dataGridViewProjects, "Projects");
                RefreshDataGrid(dataGridViewCustomers, "Customers");
                RefreshDataGrid(dataGridViewEmployees, "Employees");
                RefreshDataGrid(dataGridViewMaterials, "Materials");
                RefreshDataGrid(dataGridViewProjectMaterials, "ProjectMaterials");
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ButtonClear_Click вызывается при нажатии на кнопку "Изменить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClear_Click(object sender, EventArgs e)
        {
            try
            {
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}