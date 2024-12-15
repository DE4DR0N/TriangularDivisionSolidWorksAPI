namespace lab7
{
    partial class FormMain
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnBuild = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ColumnX = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnZ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label13 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.nmrcUpDownMargin = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownMarginX = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownMarginY = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownIters = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownFirstIter = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownStep = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP1x = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP1y = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP1z = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP7x = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP7y = new System.Windows.Forms.NumericUpDown();
            this.nmrcUpDownP7z = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMarginX)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMarginY)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownIters)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownFirstIter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownStep)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1x)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1y)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1z)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7x)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7y)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7z)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBuild
            // 
            this.btnBuild.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnBuild.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnBuild.Location = new System.Drawing.Point(829, 357);
            this.btnBuild.Margin = new System.Windows.Forms.Padding(4);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(161, 64);
            this.btnBuild.TabIndex = 0;
            this.btnBuild.Text = "Построить";
            this.btnBuild.UseVisualStyleBackColor = true;
            this.btnBuild.Click += new System.EventHandler(this.btnBuild_Click);
            // 
            // btnClear
            // 
            this.btnClear.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnClear.Enabled = false;
            this.btnClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnClear.Location = new System.Drawing.Point(998, 357);
            this.btnClear.Margin = new System.Windows.Forms.Padding(4);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(161, 64);
            this.btnClear.TabIndex = 1;
            this.btnClear.Text = "Очистить";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox1.Location = new System.Drawing.Point(622, 10);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(132, 30);
            this.textBox1.TabIndex = 2;
            // 
            // textBox2
            // 
            this.textBox2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox2.Location = new System.Drawing.Point(622, 48);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(132, 30);
            this.textBox2.TabIndex = 3;
            // 
            // textBox3
            // 
            this.textBox3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox3.Location = new System.Drawing.Point(622, 86);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(132, 30);
            this.textBox3.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(548, 14);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 24);
            this.label1.TabIndex = 8;
            this.label1.Text = "Длина";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(536, 52);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 24);
            this.label2.TabIndex = 9;
            this.label2.Text = "Ширина";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(538, 90);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(76, 24);
            this.label3.TabIndex = 10;
            this.label3.Text = "Высота";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(459, 140);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(157, 48);
            this.label5.TabIndex = 12;
            this.label5.Text = "Отступы между \r\nвырезами";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(406, 317);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(208, 24);
            this.label6.TabIndex = 13;
            this.label6.Text = "Количество итераций";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.Location = new System.Drawing.Point(829, 251);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(46, 25);
            this.label7.TabIndex = 18;
            this.label7.Text = "P1x";
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label8.Location = new System.Drawing.Point(829, 287);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(46, 25);
            this.label8.TabIndex = 20;
            this.label8.Text = "P1y";
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label9.Location = new System.Drawing.Point(829, 323);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(46, 25);
            this.label9.TabIndex = 22;
            this.label9.Text = "P1z";
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label10.Location = new System.Drawing.Point(1007, 251);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(46, 25);
            this.label10.TabIndex = 24;
            this.label10.Text = "P7x";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label11.Location = new System.Drawing.Point(1007, 287);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(46, 25);
            this.label11.TabIndex = 26;
            this.label11.Text = "P7y";
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label12.Location = new System.Drawing.Point(1007, 323);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(46, 25);
            this.label12.TabIndex = 28;
            this.label12.Text = "P7z";
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { this.ColumnX, this.ColumnY, this.ColumnZ });
            this.dataGridView1.Location = new System.Drawing.Point(772, 10);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(505, 220);
            this.dataGridView1.TabIndex = 29;
            // 
            // ColumnX
            // 
            this.ColumnX.HeaderText = "X";
            this.ColumnX.MinimumWidth = 6;
            this.ColumnX.Name = "ColumnX";
            this.ColumnX.Width = 125;
            // 
            // ColumnY
            // 
            this.ColumnY.HeaderText = "Y";
            this.ColumnY.MinimumWidth = 6;
            this.ColumnY.Name = "ColumnY";
            this.ColumnY.Width = 125;
            // 
            // ColumnZ
            // 
            this.ColumnZ.HeaderText = "Z";
            this.ColumnZ.MinimumWidth = 6;
            this.ColumnZ.Name = "ColumnZ";
            this.ColumnZ.Width = 125;
            // 
            // label13
            // 
            this.label13.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label13.Location = new System.Drawing.Point(421, 355);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(195, 24);
            this.label13.TabIndex = 31;
            this.label13.Text = "Начальная итерация";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::lab7.Properties.Resources.image1;
            this.pictureBox1.Location = new System.Drawing.Point(16, 15);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(382, 347);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 14;
            this.pictureBox1.TabStop = false;
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(571, 393);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 24);
            this.label4.TabIndex = 33;
            this.label4.Text = "Шаг";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label14
            // 
            this.label14.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label14.Location = new System.Drawing.Point(480, 188);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(134, 48);
            this.label14.TabIndex = 35;
            this.label14.Text = "Отступы по X\r\nвне фигуры";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label15.Location = new System.Drawing.Point(484, 235);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(132, 48);
            this.label15.TabIndex = 37;
            this.label15.Text = "Отступы по Y\r\nвне фигуры";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // nmrcUpDownMargin
            // 
            this.nmrcUpDownMargin.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownMargin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownMargin.Location = new System.Drawing.Point(622, 149);
            this.nmrcUpDownMargin.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.nmrcUpDownMargin.Name = "nmrcUpDownMargin";
            this.nmrcUpDownMargin.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownMargin.TabIndex = 38;
            this.nmrcUpDownMargin.Value = new decimal(new int[] { 2, 0, 0, 0 });
            // 
            // nmrcUpDownMarginX
            // 
            this.nmrcUpDownMarginX.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownMarginX.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownMarginX.Location = new System.Drawing.Point(622, 197);
            this.nmrcUpDownMarginX.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.nmrcUpDownMarginX.Name = "nmrcUpDownMarginX";
            this.nmrcUpDownMarginX.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownMarginX.TabIndex = 39;
            this.nmrcUpDownMarginX.Value = new decimal(new int[] { 4, 0, 0, 0 });
            // 
            // nmrcUpDownMarginY
            // 
            this.nmrcUpDownMarginY.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownMarginY.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownMarginY.Location = new System.Drawing.Point(622, 244);
            this.nmrcUpDownMarginY.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.nmrcUpDownMarginY.Name = "nmrcUpDownMarginY";
            this.nmrcUpDownMarginY.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownMarginY.TabIndex = 40;
            this.nmrcUpDownMarginY.Value = new decimal(new int[] { 8, 0, 0, 0 });
            // 
            // nmrcUpDownIters
            // 
            this.nmrcUpDownIters.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownIters.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownIters.Location = new System.Drawing.Point(623, 314);
            this.nmrcUpDownIters.Name = "nmrcUpDownIters";
            this.nmrcUpDownIters.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownIters.TabIndex = 41;
            this.nmrcUpDownIters.Value = new decimal(new int[] { 5, 0, 0, 0 });
            // 
            // nmrcUpDownFirstIter
            // 
            this.nmrcUpDownFirstIter.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownFirstIter.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownFirstIter.Location = new System.Drawing.Point(623, 352);
            this.nmrcUpDownFirstIter.Name = "nmrcUpDownFirstIter";
            this.nmrcUpDownFirstIter.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownFirstIter.TabIndex = 42;
            this.nmrcUpDownFirstIter.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // nmrcUpDownStep
            // 
            this.nmrcUpDownStep.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.nmrcUpDownStep.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownStep.Location = new System.Drawing.Point(623, 390);
            this.nmrcUpDownStep.Name = "nmrcUpDownStep";
            this.nmrcUpDownStep.Size = new System.Drawing.Size(132, 30);
            this.nmrcUpDownStep.TabIndex = 43;
            this.nmrcUpDownStep.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // nmrcUpDownP1x
            // 
            this.nmrcUpDownP1x.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP1x.DecimalPlaces = 1;
            this.nmrcUpDownP1x.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP1x.Location = new System.Drawing.Point(881, 249);
            this.nmrcUpDownP1x.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP1x.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP1x.Name = "nmrcUpDownP1x";
            this.nmrcUpDownP1x.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP1x.TabIndex = 44;
            this.nmrcUpDownP1x.Value = new decimal(new int[] { 325, 0, 0, 65536 });
            // 
            // nmrcUpDownP1y
            // 
            this.nmrcUpDownP1y.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP1y.DecimalPlaces = 1;
            this.nmrcUpDownP1y.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP1y.Location = new System.Drawing.Point(881, 285);
            this.nmrcUpDownP1y.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP1y.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP1y.Name = "nmrcUpDownP1y";
            this.nmrcUpDownP1y.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP1y.TabIndex = 45;
            this.nmrcUpDownP1y.Value = new decimal(new int[] { 125, 0, 0, -2147418112 });
            // 
            // nmrcUpDownP1z
            // 
            this.nmrcUpDownP1z.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP1z.DecimalPlaces = 1;
            this.nmrcUpDownP1z.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP1z.Location = new System.Drawing.Point(881, 321);
            this.nmrcUpDownP1z.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP1z.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP1z.Name = "nmrcUpDownP1z";
            this.nmrcUpDownP1z.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP1z.TabIndex = 46;
            this.nmrcUpDownP1z.Value = new decimal(new int[] { 25, 0, 0, -2147418112 });
            // 
            // nmrcUpDownP7x
            // 
            this.nmrcUpDownP7x.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP7x.DecimalPlaces = 1;
            this.nmrcUpDownP7x.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP7x.Location = new System.Drawing.Point(1059, 249);
            this.nmrcUpDownP7x.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP7x.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP7x.Name = "nmrcUpDownP7x";
            this.nmrcUpDownP7x.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP7x.TabIndex = 47;
            this.nmrcUpDownP7x.Value = new decimal(new int[] { 325, 0, 0, -2147418112 });
            // 
            // nmrcUpDownP7y
            // 
            this.nmrcUpDownP7y.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP7y.DecimalPlaces = 1;
            this.nmrcUpDownP7y.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP7y.Location = new System.Drawing.Point(1059, 285);
            this.nmrcUpDownP7y.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP7y.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP7y.Name = "nmrcUpDownP7y";
            this.nmrcUpDownP7y.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP7y.TabIndex = 48;
            this.nmrcUpDownP7y.Value = new decimal(new int[] { 325, 0, 0, 65536 });
            // 
            // nmrcUpDownP7z
            // 
            this.nmrcUpDownP7z.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nmrcUpDownP7z.DecimalPlaces = 1;
            this.nmrcUpDownP7z.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nmrcUpDownP7z.Location = new System.Drawing.Point(1059, 321);
            this.nmrcUpDownP7z.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.nmrcUpDownP7z.Minimum = new decimal(new int[] { 1000, 0, 0, -2147483648 });
            this.nmrcUpDownP7z.Name = "nmrcUpDownP7z";
            this.nmrcUpDownP7z.Size = new System.Drawing.Size(100, 30);
            this.nmrcUpDownP7z.TabIndex = 49;
            this.nmrcUpDownP7z.Value = new decimal(new int[] { 475, 0, 0, -2147418112 });
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1289, 429);
            this.Controls.Add(this.nmrcUpDownP7z);
            this.Controls.Add(this.nmrcUpDownP7y);
            this.Controls.Add(this.nmrcUpDownP7x);
            this.Controls.Add(this.nmrcUpDownP1z);
            this.Controls.Add(this.nmrcUpDownP1y);
            this.Controls.Add(this.nmrcUpDownP1x);
            this.Controls.Add(this.nmrcUpDownStep);
            this.Controls.Add(this.nmrcUpDownFirstIter);
            this.Controls.Add(this.nmrcUpDownIters);
            this.Controls.Add(this.nmrcUpDownMarginY);
            this.Controls.Add(this.nmrcUpDownMarginX);
            this.Controls.Add(this.nmrcUpDownMargin);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnBuild);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "FormMain";
            this.Text = "Лаб 7";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMarginX)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownMarginY)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownIters)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownFirstIter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownStep)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1x)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1y)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP1z)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7x)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7y)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmrcUpDownP7z)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.NumericUpDown nmrcUpDownP1y;
        private System.Windows.Forms.NumericUpDown nmrcUpDownP1z;
        private System.Windows.Forms.NumericUpDown nmrcUpDownP7x;
        private System.Windows.Forms.NumericUpDown nmrcUpDownP7y;

        private System.Windows.Forms.NumericUpDown nmrcUpDownP7z;

        private System.Windows.Forms.NumericUpDown nmrcUpDownP1x;

        private System.Windows.Forms.NumericUpDown nmrcUpDownIters;
        private System.Windows.Forms.NumericUpDown nmrcUpDownFirstIter;
        private System.Windows.Forms.NumericUpDown nmrcUpDownStep;

        private System.Windows.Forms.NumericUpDown nmrcUpDownMargin;
        private System.Windows.Forms.NumericUpDown nmrcUpDownMarginY;

        private System.Windows.Forms.NumericUpDown nmrcUpDownMarginX;

        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;

        #endregion

        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnX;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnY;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnZ;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label4;
    }
}

