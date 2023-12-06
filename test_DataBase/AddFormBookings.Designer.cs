namespace test_DataBase
{
    partial class AddFormBookings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonSave = new System.Windows.Forms.Button();
            this.label25 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.textBoxTotalAmount = new System.Windows.Forms.TextBox();
            this.textBoxNumberOfPersons = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.textBoxBookingDate = new System.Windows.Forms.TextBox();
            this.textBoxTourIDBookings = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.textBoxClientIDBookings = new System.Windows.Forms.TextBox();
            this.textBoxBookingID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(310, 54);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(189, 32);
            this.label1.TabIndex = 38;
            this.label1.Text = "Бронирование";
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelTitle.Location = new System.Drawing.Point(309, 15);
            this.labelTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(267, 40);
            this.labelTitle.TabIndex = 37;
            this.labelTitle.Text = "Создание записи:";
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(422, 1018);
            this.buttonSave.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(303, 86);
            this.buttonSave.TabIndex = 28;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(246, 850);
            this.label25.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(115, 20);
            this.label25.TabIndex = 50;
            this.label25.Text = "Общая сумма:";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(190, 790);
            this.label26.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(171, 20);
            this.label26.TabIndex = 49;
            this.label26.Text = "Количество человек:";
            // 
            // textBoxTotalAmount
            // 
            this.textBoxTotalAmount.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxTotalAmount.Location = new System.Drawing.Point(369, 831);
            this.textBoxTotalAmount.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxTotalAmount.Name = "textBoxTotalAmount";
            this.textBoxTotalAmount.Size = new System.Drawing.Size(584, 45);
            this.textBoxTotalAmount.TabIndex = 48;
            // 
            // textBoxNumberOfPersons
            // 
            this.textBoxNumberOfPersons.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxNumberOfPersons.Location = new System.Drawing.Point(369, 771);
            this.textBoxNumberOfPersons.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxNumberOfPersons.Name = "textBoxNumberOfPersons";
            this.textBoxNumberOfPersons.Size = new System.Drawing.Size(584, 45);
            this.textBoxNumberOfPersons.TabIndex = 47;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(197, 735);
            this.label23.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(164, 20);
            this.label23.TabIndex = 46;
            this.label23.Text = "Дата бронирования:";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(260, 675);
            this.label24.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(101, 20);
            this.label24.TabIndex = 45;
            this.label24.Text = "Номер тура:";
            // 
            // textBoxBookingDate
            // 
            this.textBoxBookingDate.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxBookingDate.Location = new System.Drawing.Point(369, 716);
            this.textBoxBookingDate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxBookingDate.Name = "textBoxBookingDate";
            this.textBoxBookingDate.Size = new System.Drawing.Size(584, 45);
            this.textBoxBookingDate.TabIndex = 44;
            // 
            // textBoxTourIDBookings
            // 
            this.textBoxTourIDBookings.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxTourIDBookings.Location = new System.Drawing.Point(369, 656);
            this.textBoxTourIDBookings.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxTourIDBookings.Name = "textBoxTourIDBookings";
            this.textBoxTourIDBookings.Size = new System.Drawing.Size(584, 45);
            this.textBoxTourIDBookings.TabIndex = 43;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(231, 620);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(130, 20);
            this.label7.TabIndex = 42;
            this.label7.Text = "Номер клиента:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(298, 560);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 20);
            this.label8.TabIndex = 41;
            this.label8.Text = "Номер:";
            // 
            // textBoxClientIDBookings
            // 
            this.textBoxClientIDBookings.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxClientIDBookings.Location = new System.Drawing.Point(369, 601);
            this.textBoxClientIDBookings.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxClientIDBookings.Name = "textBoxClientIDBookings";
            this.textBoxClientIDBookings.Size = new System.Drawing.Size(584, 45);
            this.textBoxClientIDBookings.TabIndex = 40;
            // 
            // textBoxBookingID
            // 
            this.textBoxBookingID.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxBookingID.Location = new System.Drawing.Point(369, 541);
            this.textBoxBookingID.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxBookingID.Name = "textBoxBookingID";
            this.textBoxBookingID.Size = new System.Drawing.Size(584, 45);
            this.textBoxBookingID.TabIndex = 39;
            // 
            // AddFormBookings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1152, 1122);
            this.Controls.Add(this.label25);
            this.Controls.Add(this.label26);
            this.Controls.Add(this.textBoxTotalAmount);
            this.Controls.Add(this.textBoxNumberOfPersons);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.label24);
            this.Controls.Add(this.textBoxBookingDate);
            this.Controls.Add(this.textBoxTourIDBookings);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBoxClientIDBookings);
            this.Controls.Add(this.textBoxBookingID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.buttonSave);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "AddFormBookings";
            this.Text = "Добавить бронирование";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.TextBox textBoxTotalAmount;
        private System.Windows.Forms.TextBox textBoxNumberOfPersons;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox textBoxBookingDate;
        private System.Windows.Forms.TextBox textBoxTourIDBookings;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBoxClientIDBookings;
        private System.Windows.Forms.TextBox textBoxBookingID;
    }
}