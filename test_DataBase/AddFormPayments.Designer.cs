namespace test_DataBase
{
    partial class AddFormPayments
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
            this.label27 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.textBoxAmount = new System.Windows.Forms.TextBox();
            this.textBoxPaymentDate = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.textBoxBookingIDPayments = new System.Windows.Forms.TextBox();
            this.textBoxPaymentID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(310, 54);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(103, 32);
            this.label1.TabIndex = 38;
            this.label1.Text = "Платеж";
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
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(298, 730);
            this.label27.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(62, 20);
            this.label27.TabIndex = 46;
            this.label27.Text = "Сумма:";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(238, 675);
            this.label28.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(122, 20);
            this.label28.TabIndex = 45;
            this.label28.Text = "Дата платежа:";
            // 
            // textBoxAmount
            // 
            this.textBoxAmount.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxAmount.Location = new System.Drawing.Point(368, 716);
            this.textBoxAmount.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxAmount.Name = "textBoxAmount";
            this.textBoxAmount.Size = new System.Drawing.Size(584, 45);
            this.textBoxAmount.TabIndex = 44;
            // 
            // textBoxPaymentDate
            // 
            this.textBoxPaymentDate.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxPaymentDate.Location = new System.Drawing.Point(368, 656);
            this.textBoxPaymentDate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxPaymentDate.Name = "textBoxPaymentDate";
            this.textBoxPaymentDate.Size = new System.Drawing.Size(584, 45);
            this.textBoxPaymentDate.TabIndex = 43;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(185, 620);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(175, 20);
            this.label12.TabIndex = 42;
            this.label12.Text = "Номер бронирования:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(297, 560);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(63, 20);
            this.label13.TabIndex = 41;
            this.label13.Text = "Номер:";
            // 
            // textBoxBookingIDPayments
            // 
            this.textBoxBookingIDPayments.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxBookingIDPayments.Location = new System.Drawing.Point(368, 601);
            this.textBoxBookingIDPayments.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxBookingIDPayments.Name = "textBoxBookingIDPayments";
            this.textBoxBookingIDPayments.Size = new System.Drawing.Size(584, 45);
            this.textBoxBookingIDPayments.TabIndex = 40;
            // 
            // textBoxPaymentID
            // 
            this.textBoxPaymentID.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxPaymentID.Location = new System.Drawing.Point(368, 541);
            this.textBoxPaymentID.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxPaymentID.Name = "textBoxPaymentID";
            this.textBoxPaymentID.Size = new System.Drawing.Size(584, 45);
            this.textBoxPaymentID.TabIndex = 39;
            // 
            // AddFormPayments
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1152, 1122);
            this.Controls.Add(this.label27);
            this.Controls.Add(this.label28);
            this.Controls.Add(this.textBoxAmount);
            this.Controls.Add(this.textBoxPaymentDate);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.textBoxBookingIDPayments);
            this.Controls.Add(this.textBoxPaymentID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.buttonSave);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "AddFormPayments";
            this.Text = "Добавить платеж";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox textBoxAmount;
        private System.Windows.Forms.TextBox textBoxPaymentDate;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox textBoxBookingIDPayments;
        private System.Windows.Forms.TextBox textBoxPaymentID;
    }
}