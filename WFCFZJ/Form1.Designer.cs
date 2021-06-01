namespace GrabSteel
{
    partial class form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnZQ = new System.Windows.Forms.Button();
            this.btxExe = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.mytimer = new System.Windows.Forms.Timer(this.components);
            this.ck_Auto = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.pager1 = new DotNet.WinFormPager.PageNavigator.Pager();
            this.编号 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.标题 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.链接 = new System.Windows.Forms.DataGridViewLinkColumn();
            this.文章发布时间 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.抓取时间 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnZQ
            // 
            this.btnZQ.Location = new System.Drawing.Point(918, 35);
            this.btnZQ.Name = "btnZQ";
            this.btnZQ.Size = new System.Drawing.Size(75, 23);
            this.btnZQ.TabIndex = 0;
            this.btnZQ.Text = "全部抓取";
            this.btnZQ.UseVisualStyleBackColor = true;
            this.btnZQ.Click += new System.EventHandler(this.btnZQ_Click);
            // 
            // btxExe
            // 
            this.btxExe.Location = new System.Drawing.Point(999, 35);
            this.btxExe.Name = "btxExe";
            this.btxExe.Size = new System.Drawing.Size(75, 23);
            this.btxExe.TabIndex = 2;
            this.btxExe.Text = "导出Execl";
            this.btxExe.UseVisualStyleBackColor = true;
            this.btxExe.Click += new System.EventHandler(this.btxExe_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(100, 37);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(131, 21);
            this.textBox1.TabIndex = 5;
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(237, 37);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 6;
            this.btnSearch.Text = "查询";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(65, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "标题";
            // 
            // mytimer
            // 
            this.mytimer.Interval = 500;
            this.mytimer.Tick += new System.EventHandler(this.mytimer_Tick);
            // 
            // ck_Auto
            // 
            this.ck_Auto.AutoSize = true;
            this.ck_Auto.Checked = true;
            this.ck_Auto.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ck_Auto.Location = new System.Drawing.Point(834, 39);
            this.ck_Auto.Name = "ck_Auto";
            this.ck_Auto.Size = new System.Drawing.Size(72, 16);
            this.ck_Auto.TabIndex = 8;
            this.ck_Auto.Text = "自动抓取";
            this.ck_Auto.UseVisualStyleBackColor = true;
            this.ck_Auto.CheckedChanged += new System.EventHandler(this.ck_Auto_CheckedChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.编号,
            this.标题,
            this.链接,
            this.文章发布时间,
            this.抓取时间});
            this.dataGridView1.Location = new System.Drawing.Point(31, 81);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(1043, 383);
            this.dataGridView1.TabIndex = 9;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // pager1
            // 
            this.pager1.Location = new System.Drawing.Point(31, 480);
            this.pager1.Name = "pager1";
            this.pager1.PageIndex = 0;
            this.pager1.PageSize = 0;
            this.pager1.RecordCount = 0;
            this.pager1.Size = new System.Drawing.Size(1043, 26);
            this.pager1.TabIndex = 10;
            this.pager1.OnPageChanged += new System.EventHandler(this.pager1_OnPageChanged);
            // 
            // 编号
            // 
            this.编号.DataPropertyName = "grabID";
            this.编号.HeaderText = "编号";
            this.编号.Name = "编号";
            this.编号.Visible = false;
            // 
            // 标题
            // 
            this.标题.DataPropertyName = "title";
            this.标题.HeaderText = "标题";
            this.标题.Name = "标题";
            this.标题.Width = 300;
            // 
            // 链接
            // 
            this.链接.DataPropertyName = "url";
            this.链接.HeaderText = "链接";
            this.链接.Name = "链接";
            this.链接.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.链接.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.链接.Width = 370;
            // 
            // 文章发布时间
            // 
            this.文章发布时间.DataPropertyName = "webdate";
            this.文章发布时间.HeaderText = "文章发布时间";
            this.文章发布时间.Name = "文章发布时间";
            this.文章发布时间.Width = 160;
            // 
            // 抓取时间
            // 
            this.抓取时间.DataPropertyName = "createTime";
            this.抓取时间.HeaderText = "抓取时间";
            this.抓取时间.Name = "抓取时间";
            this.抓取时间.Width = 160;
            // 
            // form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1101, 529);
            this.Controls.Add(this.pager1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.ck_Auto);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btxExe);
            this.Controls.Add(this.btnZQ);
            this.Name = "form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "我的钢铁网-信息抓取";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.form1_FormClosed);
            this.Load += new System.EventHandler(this.form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnZQ;
        private System.Windows.Forms.Button btxExe;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer mytimer;
        private System.Windows.Forms.CheckBox ck_Auto;
        private System.Windows.Forms.DataGridView dataGridView1;
        private DotNet.WinFormPager.PageNavigator.Pager pager1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 编号;
        private System.Windows.Forms.DataGridViewTextBoxColumn 标题;
        private System.Windows.Forms.DataGridViewLinkColumn 链接;
        private System.Windows.Forms.DataGridViewTextBoxColumn 文章发布时间;
        private System.Windows.Forms.DataGridViewTextBoxColumn 抓取时间;
    }
}

