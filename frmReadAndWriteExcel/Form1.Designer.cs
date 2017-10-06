namespace frmReadAndWriteExcel
{
    partial class FrmReadAndSaveExcel
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
            this.LsvLoadData = new System.Windows.Forms.ListView();
            this.BtnLoad = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // LsvLoadData
            // 
            this.LsvLoadData.GridLines = true;
            this.LsvLoadData.Location = new System.Drawing.Point(12, 12);
            this.LsvLoadData.MultiSelect = false;
            this.LsvLoadData.Name = "LsvLoadData";
            this.LsvLoadData.Size = new System.Drawing.Size(387, 212);
            this.LsvLoadData.TabIndex = 0;
            this.LsvLoadData.UseCompatibleStateImageBehavior = false;
            this.LsvLoadData.View = System.Windows.Forms.View.Details;
            // 
            // BtnLoad
            // 
            this.BtnLoad.Location = new System.Drawing.Point(330, 250);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(75, 23);
            this.BtnLoad.TabIndex = 1;
            this.BtnLoad.Text = "&Load";
            this.BtnLoad.UseVisualStyleBackColor = true;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(221, 250);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(75, 23);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "&Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // FrmReadAndSaveExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(411, 291);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnLoad);
            this.Controls.Add(this.LsvLoadData);
            this.Name = "FrmReadAndSaveExcel";
            this.Text = "ReadAndSaveExcel";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView LsvLoadData;
        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.Button BtnSave;
    }
}

