namespace Spectro_AS7262
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series7 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series8 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series9 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series10 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series11 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series12 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title2 = new System.Windows.Forms.DataVisualization.Charting.Title();
            this.cbxPorts = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSndSettings = new System.Windows.Forms.Button();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.cbxGain = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.btnSoftReset = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxLBC = new System.Windows.Forms.ComboBox();
            this.cbxLIC = new System.Windows.Forms.ComboBox();
            this.cbxLBE = new System.Windows.Forms.ComboBox();
            this.cbxLIE = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbxMMode = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.btnSave = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.cbxCh1 = new System.Windows.Forms.CheckBox();
            this.cbxCh2 = new System.Windows.Forms.CheckBox();
            this.cbxCh3 = new System.Windows.Forms.CheckBox();
            this.cbxCh4 = new System.Windows.Forms.CheckBox();
            this.cbxCh5 = new System.Windows.Forms.CheckBox();
            this.cbxCh6 = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBlue = new System.Windows.Forms.CheckBox();
            this.checkGreen = new System.Windows.Forms.CheckBox();
            this.checkRed = new System.Windows.Forms.CheckBox();
            this.btnSndLed = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // cbxPorts
            // 
            this.cbxPorts.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxPorts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxPorts.FormattingEnabled = true;
            this.cbxPorts.Location = new System.Drawing.Point(78, 17);
            this.cbxPorts.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.cbxPorts.Name = "cbxPorts";
            this.cbxPorts.Size = new System.Drawing.Size(364, 28);
            this.cbxPorts.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(15, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 20);
            this.label1.TabIndex = 10;
            this.label1.Text = "Ports";
            // 
            // btnConnect
            // 
            this.btnConnect.BackColor = System.Drawing.SystemColors.Control;
            this.btnConnect.Location = new System.Drawing.Point(271, 61);
            this.btnConnect.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(171, 36);
            this.btnConnect.TabIndex = 8;
            this.btnConnect.Text = "Connect";
            this.btnConnect.UseVisualStyleBackColor = false;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.SystemColors.Control;
            this.btnRefresh.Location = new System.Drawing.Point(15, 61);
            this.btnRefresh.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(171, 36);
            this.btnRefresh.TabIndex = 9;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // serialPort1
            // 
            this.serialPort1.BaudRate = 115200;
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.SystemColors.Window;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 684);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(2, 0, 21, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1130, 22);
            this.statusStrip1.TabIndex = 12;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(39, 17);
            this.lblStatus.Text = "Status";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSndSettings);
            this.groupBox1.Controls.Add(this.numericUpDown1);
            this.groupBox1.Controls.Add(this.cbxGain);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.btnSoftReset);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cbxLBC);
            this.groupBox1.Controls.Add(this.cbxLIC);
            this.groupBox1.Controls.Add(this.cbxLBE);
            this.groupBox1.Controls.Add(this.cbxLIE);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cbxMMode);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(17, 186);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(427, 310);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Settings";
            // 
            // btnSndSettings
            // 
            this.btnSndSettings.BackColor = System.Drawing.SystemColors.Control;
            this.btnSndSettings.Enabled = false;
            this.btnSndSettings.Location = new System.Drawing.Point(14, 218);
            this.btnSndSettings.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSndSettings.Name = "btnSndSettings";
            this.btnSndSettings.Size = new System.Drawing.Size(398, 35);
            this.btnSndSettings.TabIndex = 6;
            this.btnSndSettings.Text = "Send Settings";
            this.btnSndSettings.UseVisualStyleBackColor = false;
            this.btnSndSettings.Click += new System.EventHandler(this.btnSndSettings_Click);
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(341, 178);
            this.numericUpDown1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(71, 26);
            this.numericUpDown1.TabIndex = 5;
            // 
            // cbxGain
            // 
            this.cbxGain.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxGain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxGain.FormattingEnabled = true;
            this.cbxGain.Items.AddRange(new object[] {
            "1x",
            "3.7x",
            "16x",
            "64x"});
            this.cbxGain.Location = new System.Drawing.Point(122, 177);
            this.cbxGain.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxGain.Name = "cbxGain";
            this.cbxGain.Size = new System.Drawing.Size(79, 28);
            this.cbxGain.TabIndex = 4;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 182);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 20);
            this.label7.TabIndex = 3;
            this.label7.Text = "Gain";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(209, 182);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(124, 20);
            this.label8.TabIndex = 2;
            this.label8.Text = "Integration Time";
            // 
            // btnSoftReset
            // 
            this.btnSoftReset.BackColor = System.Drawing.SystemColors.Control;
            this.btnSoftReset.Enabled = false;
            this.btnSoftReset.Location = new System.Drawing.Point(14, 263);
            this.btnSoftReset.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSoftReset.Name = "btnSoftReset";
            this.btnSoftReset.Size = new System.Drawing.Size(398, 35);
            this.btnSoftReset.TabIndex = 6;
            this.btnSoftReset.Text = "Soft Reset";
            this.btnSoftReset.UseVisualStyleBackColor = false;
            this.btnSoftReset.Click += new System.EventHandler(this.btnSoftReset_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(209, 140);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(62, 20);
            this.label6.TabIndex = 2;
            this.label6.Text = "Current";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(209, 98);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(62, 20);
            this.label4.TabIndex = 2;
            this.label4.Text = "Current";
            // 
            // cbxLBC
            // 
            this.cbxLBC.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxLBC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxLBC.FormattingEnabled = true;
            this.cbxLBC.Items.AddRange(new object[] {
            "12.5mA",
            "25mA",
            "50mA",
            "100mA"});
            this.cbxLBC.Location = new System.Drawing.Point(279, 135);
            this.cbxLBC.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxLBC.Name = "cbxLBC";
            this.cbxLBC.Size = new System.Drawing.Size(133, 28);
            this.cbxLBC.TabIndex = 1;
            // 
            // cbxLIC
            // 
            this.cbxLIC.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxLIC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxLIC.FormattingEnabled = true;
            this.cbxLIC.Items.AddRange(new object[] {
            "1mA",
            "2mA",
            "4mA",
            "8mA"});
            this.cbxLIC.Location = new System.Drawing.Point(279, 94);
            this.cbxLIC.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxLIC.Name = "cbxLIC";
            this.cbxLIC.Size = new System.Drawing.Size(133, 28);
            this.cbxLIC.TabIndex = 1;
            // 
            // cbxLBE
            // 
            this.cbxLBE.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxLBE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxLBE.FormattingEnabled = true;
            this.cbxLBE.Items.AddRange(new object[] {
            "ON",
            "OFF"});
            this.cbxLBE.Location = new System.Drawing.Point(122, 135);
            this.cbxLBE.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxLBE.Name = "cbxLBE";
            this.cbxLBE.Size = new System.Drawing.Size(79, 28);
            this.cbxLBE.TabIndex = 1;
            // 
            // cbxLIE
            // 
            this.cbxLIE.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxLIE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxLIE.FormattingEnabled = true;
            this.cbxLIE.Items.AddRange(new object[] {
            "ON",
            "OFF"});
            this.cbxLIE.Location = new System.Drawing.Point(122, 94);
            this.cbxLIE.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxLIE.Name = "cbxLIE";
            this.cbxLIE.Size = new System.Drawing.Size(79, 28);
            this.cbxLIE.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 140);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 20);
            this.label5.TabIndex = 0;
            this.label5.Text = "Led Bulb";
            // 
            // cbxMMode
            // 
            this.cbxMMode.BackColor = System.Drawing.Color.GhostWhite;
            this.cbxMMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxMMode.FormattingEnabled = true;
            this.cbxMMode.Items.AddRange(new object[] {
            "Mode 0",
            "Mode 1",
            "Mode 2",
            "Mode 3"});
            this.cbxMMode.Location = new System.Drawing.Point(168, 35);
            this.cbxMMode.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cbxMMode.Name = "cbxMMode";
            this.cbxMMode.Size = new System.Drawing.Size(244, 28);
            this.cbxMMode.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 98);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(102, 20);
            this.label3.TabIndex = 0;
            this.label3.Text = "Led Indicator";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 40);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(151, 20);
            this.label2.TabIndex = 0;
            this.label2.Text = "Measurement mode";
            // 
            // timer1
            // 
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.SystemColors.Control;
            this.btnSave.Location = new System.Drawing.Point(240, 611);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(204, 35);
            this.btnSave.TabIndex = 18;
            this.btnSave.Text = "Save Data";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.SystemColors.Control;
            this.btnStart.Location = new System.Drawing.Point(14, 566);
            this.btnStart.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(427, 35);
            this.btnStart.TabIndex = 19;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnClear
            // 
            this.btnClear.BackColor = System.Drawing.SystemColors.Control;
            this.btnClear.Location = new System.Drawing.Point(14, 611);
            this.btnClear.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(204, 35);
            this.btnClear.TabIndex = 20;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = false;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(452, 660);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(33, 34);
            this.textBox1.TabIndex = 21;
            this.textBox1.Visible = false;
            // 
            // chart1
            // 
            this.chart1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            chartArea2.AxisX.MajorGrid.LineColor = System.Drawing.Color.LightGray;
            chartArea2.AxisX.Maximum = 100D;
            chartArea2.AxisX.MinorGrid.LineColor = System.Drawing.Color.Gray;
            chartArea2.AxisX.Title = "Samples";
            chartArea2.AxisX.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            chartArea2.AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray;
            chartArea2.AxisY.Title = "Amplitude";
            chartArea2.AxisY.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            chartArea2.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea2);
            legend2.Alignment = System.Drawing.StringAlignment.Center;
            legend2.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom;
            legend2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            legend2.IsTextAutoFit = false;
            legend2.Name = "Legend1";
            this.chart1.Legends.Add(legend2);
            this.chart1.Location = new System.Drawing.Point(452, 14);
            this.chart1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chart1.Name = "chart1";
            series7.ChartArea = "ChartArea1";
            series7.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series7.Color = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            series7.Legend = "Legend1";
            series7.Name = "Channel 1";
            series8.ChartArea = "ChartArea1";
            series8.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series8.Color = System.Drawing.Color.Blue;
            series8.Legend = "Legend1";
            series8.Name = "Channel 2";
            series9.ChartArea = "ChartArea1";
            series9.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series9.Color = System.Drawing.Color.Green;
            series9.Legend = "Legend1";
            series9.Name = "Channel 3";
            series10.ChartArea = "ChartArea1";
            series10.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series10.Color = System.Drawing.Color.Goldenrod;
            series10.Legend = "Legend1";
            series10.Name = "Channel 4";
            series11.ChartArea = "ChartArea1";
            series11.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series11.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            series11.Legend = "Legend1";
            series11.Name = "Channel 5";
            series12.ChartArea = "ChartArea1";
            series12.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series12.Color = System.Drawing.Color.Red;
            series12.Legend = "Legend1";
            series12.Name = "Channel 6";
            this.chart1.Series.Add(series7);
            this.chart1.Series.Add(series8);
            this.chart1.Series.Add(series9);
            this.chart1.Series.Add(series10);
            this.chart1.Series.Add(series11);
            this.chart1.Series.Add(series12);
            this.chart1.Size = new System.Drawing.Size(665, 620);
            this.chart1.TabIndex = 22;
            this.chart1.Text = "chart1";
            title2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            title2.Name = "Title1";
            title2.Text = "Sensor Value";
            this.chart1.Titles.Add(title2);
            // 
            // cbxCh1
            // 
            this.cbxCh1.AutoSize = true;
            this.cbxCh1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.cbxCh1.Location = new System.Drawing.Point(30, 504);
            this.cbxCh1.Name = "cbxCh1";
            this.cbxCh1.Size = new System.Drawing.Size(124, 24);
            this.cbxCh1.TabIndex = 7;
            this.cbxCh1.Text = "Ch 1 (450nm)";
            this.cbxCh1.UseVisualStyleBackColor = true;
            // 
            // cbxCh2
            // 
            this.cbxCh2.AutoSize = true;
            this.cbxCh2.ForeColor = System.Drawing.Color.Blue;
            this.cbxCh2.Location = new System.Drawing.Point(164, 504);
            this.cbxCh2.Name = "cbxCh2";
            this.cbxCh2.Size = new System.Drawing.Size(124, 24);
            this.cbxCh2.TabIndex = 7;
            this.cbxCh2.Text = "Ch 2 (500nm)";
            this.cbxCh2.UseVisualStyleBackColor = true;
            // 
            // cbxCh3
            // 
            this.cbxCh3.AutoSize = true;
            this.cbxCh3.ForeColor = System.Drawing.Color.Green;
            this.cbxCh3.Location = new System.Drawing.Point(305, 504);
            this.cbxCh3.Name = "cbxCh3";
            this.cbxCh3.Size = new System.Drawing.Size(124, 24);
            this.cbxCh3.TabIndex = 7;
            this.cbxCh3.Text = "Ch 3 (550nm)";
            this.cbxCh3.UseVisualStyleBackColor = true;
            // 
            // cbxCh4
            // 
            this.cbxCh4.AutoSize = true;
            this.cbxCh4.ForeColor = System.Drawing.Color.Goldenrod;
            this.cbxCh4.Location = new System.Drawing.Point(30, 534);
            this.cbxCh4.Name = "cbxCh4";
            this.cbxCh4.Size = new System.Drawing.Size(124, 24);
            this.cbxCh4.TabIndex = 7;
            this.cbxCh4.Text = "Ch 4 (570nm)";
            this.cbxCh4.UseVisualStyleBackColor = true;
            // 
            // cbxCh5
            // 
            this.cbxCh5.AutoSize = true;
            this.cbxCh5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.cbxCh5.Location = new System.Drawing.Point(164, 534);
            this.cbxCh5.Name = "cbxCh5";
            this.cbxCh5.Size = new System.Drawing.Size(124, 24);
            this.cbxCh5.TabIndex = 7;
            this.cbxCh5.Text = "Ch 5 (600nm)";
            this.cbxCh5.UseVisualStyleBackColor = true;
            // 
            // cbxCh6
            // 
            this.cbxCh6.AutoSize = true;
            this.cbxCh6.ForeColor = System.Drawing.Color.Red;
            this.cbxCh6.Location = new System.Drawing.Point(305, 534);
            this.cbxCh6.Name = "cbxCh6";
            this.cbxCh6.Size = new System.Drawing.Size(124, 24);
            this.cbxCh6.TabIndex = 7;
            this.cbxCh6.Text = "Ch 6 (650nm)";
            this.cbxCh6.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.Location = new System.Drawing.Point(452, 644);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(665, 35);
            this.button1.TabIndex = 18;
            this.button1.Text = "Export Image";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBlue
            // 
            this.checkBlue.AutoSize = true;
            this.checkBlue.ForeColor = System.Drawing.Color.Blue;
            this.checkBlue.Location = new System.Drawing.Point(353, 110);
            this.checkBlue.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.checkBlue.Name = "checkBlue";
            this.checkBlue.Size = new System.Drawing.Size(91, 24);
            this.checkBlue.TabIndex = 24;
            this.checkBlue.Text = "Blue Led";
            this.checkBlue.UseVisualStyleBackColor = true;
            // 
            // checkGreen
            // 
            this.checkGreen.AutoSize = true;
            this.checkGreen.ForeColor = System.Drawing.Color.Green;
            this.checkGreen.Location = new System.Drawing.Point(177, 108);
            this.checkGreen.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.checkGreen.Name = "checkGreen";
            this.checkGreen.Size = new System.Drawing.Size(104, 24);
            this.checkGreen.TabIndex = 25;
            this.checkGreen.Text = "Green Led";
            this.checkGreen.UseVisualStyleBackColor = true;
            // 
            // checkRed
            // 
            this.checkRed.AutoSize = true;
            this.checkRed.ForeColor = System.Drawing.Color.Red;
            this.checkRed.Location = new System.Drawing.Point(15, 110);
            this.checkRed.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.checkRed.Name = "checkRed";
            this.checkRed.Size = new System.Drawing.Size(89, 24);
            this.checkRed.TabIndex = 26;
            this.checkRed.Text = "Red Led";
            this.checkRed.UseVisualStyleBackColor = true;
            // 
            // btnSndLed
            // 
            this.btnSndLed.BackColor = System.Drawing.SystemColors.Control;
            this.btnSndLed.Enabled = false;
            this.btnSndLed.Location = new System.Drawing.Point(15, 141);
            this.btnSndLed.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSndLed.Name = "btnSndLed";
            this.btnSndLed.Size = new System.Drawing.Size(429, 35);
            this.btnSndLed.TabIndex = 23;
            this.btnSndLed.Text = "Change LED color\r\n";
            this.btnSndLed.UseVisualStyleBackColor = false;
            this.btnSndLed.Click += new System.EventHandler(this.btnSndLed_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1130, 706);
            this.Controls.Add(this.checkBlue);
            this.Controls.Add(this.checkGreen);
            this.Controls.Add(this.checkRed);
            this.Controls.Add(this.btnSndLed);
            this.Controls.Add(this.cbxCh6);
            this.Controls.Add(this.cbxCh3);
            this.Controls.Add(this.cbxCh5);
            this.Controls.Add(this.cbxCh2);
            this.Controls.Add(this.cbxCh4);
            this.Controls.Add(this.cbxCh1);
            this.Controls.Add(this.chart1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.cbxPorts);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.btnRefresh);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.Text = "Spectro Photometer AS7262";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxPorts;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnRefresh;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbxLIC;
        private System.Windows.Forms.ComboBox cbxLIE;
        private System.Windows.Forms.ComboBox cbxMMode;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cbxLBC;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.ComboBox cbxGain;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbxLBE;
        private System.Windows.Forms.Button btnSndSettings;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.Button btnSoftReset;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.CheckBox cbxCh1;
        private System.Windows.Forms.CheckBox cbxCh2;
        private System.Windows.Forms.CheckBox cbxCh3;
        private System.Windows.Forms.CheckBox cbxCh4;
        private System.Windows.Forms.CheckBox cbxCh5;
        private System.Windows.Forms.CheckBox cbxCh6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBlue;
        private System.Windows.Forms.CheckBox checkGreen;
        private System.Windows.Forms.CheckBox checkRed;
        private System.Windows.Forms.Button btnSndLed;
    }
}

