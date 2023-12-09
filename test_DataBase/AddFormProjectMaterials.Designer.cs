namespace test_DataBase
{
    partial class AddFormProjectMaterials
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
            this.label31 = new System.Windows.Forms.Label();
            this.textBoxQuantinityUsed = new System.Windows.Forms.TextBox();
            this.label32 = new System.Windows.Forms.Label();
            this.label33 = new System.Windows.Forms.Label();
            this.textBoxMaterialIDProjectMaterials = new System.Windows.Forms.TextBox();
            this.textBoxProjectIDProjectMaterials = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(210, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 21);
            this.label1.TabIndex = 32;
            this.label1.Text = "Использование материала";
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelTitle.Location = new System.Drawing.Point(209, 9);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(175, 25);
            this.labelTitle.TabIndex = 31;
            this.labelTitle.Text = "Создание записи:";
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(284, 661);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(202, 56);
            this.buttonSave.TabIndex = 30;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(156, 440);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(84, 13);
            this.label31.TabIndex = 40;
            this.label31.Text = "Использовано:";
            // 
            // textBoxQuantinityUsed
            // 
            this.textBoxQuantinityUsed.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxQuantinityUsed.Location = new System.Drawing.Point(246, 428);
            this.textBoxQuantinityUsed.Name = "textBoxQuantinityUsed";
            this.textBoxQuantinityUsed.Size = new System.Drawing.Size(391, 33);
            this.textBoxQuantinityUsed.TabIndex = 39;
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Location = new System.Drawing.Point(138, 401);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(102, 13);
            this.label32.TabIndex = 38;
            this.label32.Text = "Номер материала:";
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Location = new System.Drawing.Point(152, 362);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(88, 13);
            this.label33.TabIndex = 37;
            this.label33.Text = "Номер проекта:";
            // 
            // textBoxMaterialIDProjectMaterials
            // 
            this.textBoxMaterialIDProjectMaterials.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxMaterialIDProjectMaterials.Location = new System.Drawing.Point(246, 389);
            this.textBoxMaterialIDProjectMaterials.Name = "textBoxMaterialIDProjectMaterials";
            this.textBoxMaterialIDProjectMaterials.Size = new System.Drawing.Size(391, 33);
            this.textBoxMaterialIDProjectMaterials.TabIndex = 36;
            // 
            // textBoxProjectIDProjectMaterials
            // 
            this.textBoxProjectIDProjectMaterials.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxProjectIDProjectMaterials.Location = new System.Drawing.Point(246, 350);
            this.textBoxProjectIDProjectMaterials.Name = "textBoxProjectIDProjectMaterials";
            this.textBoxProjectIDProjectMaterials.Size = new System.Drawing.Size(391, 33);
            this.textBoxProjectIDProjectMaterials.TabIndex = 35;
            // 
            // AddFormsProjectMaterials
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 729);
            this.Controls.Add(this.label31);
            this.Controls.Add(this.textBoxQuantinityUsed);
            this.Controls.Add(this.label32);
            this.Controls.Add(this.label33);
            this.Controls.Add(this.textBoxMaterialIDProjectMaterials);
            this.Controls.Add(this.textBoxProjectIDProjectMaterials);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.buttonSave);
            this.Name = "AddFormsProjectMaterials";
            this.Text = "Добавить использование материала";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.TextBox textBoxQuantinityUsed;
        private System.Windows.Forms.Label label32;
        private System.Windows.Forms.Label label33;
        private System.Windows.Forms.TextBox textBoxMaterialIDProjectMaterials;
        private System.Windows.Forms.TextBox textBoxProjectIDProjectMaterials;
    }
}