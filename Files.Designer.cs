namespace OfficeAutomation
{
    partial class Files
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Files));
            this.gvFiles = new System.Windows.Forms.DataGridView();
            this.lbFolder = new System.Windows.Forms.Label();
            this.txtFacFilter = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.gvFiles)).BeginInit();
            this.SuspendLayout();
            // 
            // gvFiles
            // 
            this.gvFiles.AllowUserToAddRows = false;
            this.gvFiles.AllowUserToDeleteRows = false;
            this.gvFiles.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.gvFiles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvFiles.Location = new System.Drawing.Point(12, 64);
            this.gvFiles.Name = "gvFiles";
            this.gvFiles.Size = new System.Drawing.Size(319, 181);
            this.gvFiles.TabIndex = 0;
            this.gvFiles.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvFiles_RowHeaderMouseClick);
            // 
            // lbFolder
            // 
            this.lbFolder.AutoSize = true;
            this.lbFolder.Location = new System.Drawing.Point(20, 9);
            this.lbFolder.Name = "lbFolder";
            this.lbFolder.Size = new System.Drawing.Size(35, 13);
            this.lbFolder.TabIndex = 1;
            this.lbFolder.Text = "label1";
            // 
            // txtFacFilter
            // 
            this.txtFacFilter.Location = new System.Drawing.Point(58, 38);
            this.txtFacFilter.Name = "txtFacFilter";
            this.txtFacFilter.Size = new System.Drawing.Size(138, 20);
            this.txtFacFilter.TabIndex = 3;
            this.txtFacFilter.TextChanged += new System.EventHandler(this.txtFacFilter_TextChanged);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(20, 42);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(35, 13);
            this.label14.TabIndex = 4;
            this.label14.Text = "Filter";
            // 
            // Files
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 268);
            this.Controls.Add(this.txtFacFilter);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.lbFolder);
            this.Controls.Add(this.gvFiles);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Files";
            this.Text = "Files";
            this.Load += new System.EventHandler(this.Files_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gvFiles)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView gvFiles;
        private System.Windows.Forms.Label lbFolder;
        private System.Windows.Forms.TextBox txtFacFilter;
        private System.Windows.Forms.Label label14;
    }
}