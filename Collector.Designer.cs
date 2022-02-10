
namespace DataCollector
{
    partial class DataCollector
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.generate = new System.Windows.Forms.Button();
            this.grid = new System.Windows.Forms.DataGridView();
            this.update = new System.Windows.Forms.Button();
            this.download = new System.Windows.Forms.Button();
            this.grid_case = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grid_datas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // generate
            // 
            this.generate.Font = new System.Drawing.Font("新細明體", 12F);
            this.generate.Location = new System.Drawing.Point(12, 12);
            this.generate.Name = "generate";
            this.generate.Size = new System.Drawing.Size(776, 31);
            this.generate.TabIndex = 0;
            this.generate.Text = "存取資料";
            this.generate.UseVisualStyleBackColor = true;
            this.generate.Click += new System.EventHandler(this.generate_Click);
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.grid.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.grid.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.grid_case,
            this.grid_datas});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("新細明體", 16F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grid.DefaultCellStyle = dataGridViewCellStyle3;
            this.grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.grid.Location = new System.Drawing.Point(8, 49);
            this.grid.Name = "grid";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("新細明體", 24F);
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grid.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.grid.RowsDefaultCellStyle = dataGridViewCellStyle5;
            this.grid.RowTemplate.Height = 24;
            this.grid.Size = new System.Drawing.Size(780, 352);
            this.grid.TabIndex = 1;
            this.grid.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grid_CellContentClick);
            // 
            // update
            // 
            this.update.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.update.Font = new System.Drawing.Font("新細明體", 12F);
            this.update.Location = new System.Drawing.Point(407, 407);
            this.update.Name = "update";
            this.update.Size = new System.Drawing.Size(377, 31);
            this.update.TabIndex = 4;
            this.update.Text = "更新個案/試算表ID";
            this.update.UseVisualStyleBackColor = true;
            this.update.Click += new System.EventHandler(this.update_Click);
            // 
            // download
            // 
            this.download.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.download.Font = new System.Drawing.Font("新細明體", 12F);
            this.download.Location = new System.Drawing.Point(8, 407);
            this.download.Name = "download";
            this.download.Size = new System.Drawing.Size(380, 31);
            this.download.TabIndex = 5;
            this.download.Text = "下載試算表";
            this.download.UseVisualStyleBackColor = true;
            this.download.Click += new System.EventHandler(this.download_Click);
            // 
            // grid_case
            // 
            dataGridViewCellStyle2.Font = new System.Drawing.Font("新細明體", 16F);
            this.grid_case.DefaultCellStyle = dataGridViewCellStyle2;
            this.grid_case.HeaderText = "個案";
            this.grid_case.MinimumWidth = 10;
            this.grid_case.Name = "grid_case";
            this.grid_case.Width = 400;
            // 
            // grid_datas
            // 
            this.grid_datas.HeaderText = "填寫時間";
            this.grid_datas.Name = "grid_datas";
            this.grid_datas.Width = 332;
            // 
            // DataCollector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.download);
            this.Controls.Add(this.update);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.generate);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "DataCollector";
            this.ShowIcon = false;
            this.Text = "個案資料統整器";
            this.Load += new System.EventHandler(this.DataCollector_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button generate;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Button update;
        private System.Windows.Forms.Button download;
        private System.Windows.Forms.DataGridViewTextBoxColumn grid_case;
        private System.Windows.Forms.DataGridViewTextBoxColumn grid_datas;
    }
}

