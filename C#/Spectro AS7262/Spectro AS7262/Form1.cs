using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Spectro_AS7262
{
    public partial class Form1 : Form
    {
        private bool readEnabled;
        private string line;
        private double FactorCorrection = 5.47715;
        private StreamWriter sw;
        List<double> mean1 = new List<double>();
        List<double> mean2 = new List<double>();
        List<double> mean3 = new List<double>();
        List<double> mean4 = new List<double>();
        List<double> mean5 = new List<double>();
        List<double> mean6 = new List<double>();

        double mMean1, mMean2, mMean3, mMean4, mMean5, mMean6 = 0;
        private int lMean = 10;

        double[] reference = { 0, 0, 0, 0, 0, 0 };

        public Form1()
        {
            InitializeComponent();
            refreshPorts();
            foreach (Control item in groupBox1.Controls)
            {
                if (item is ComboBox)
                    ((ComboBox)item).SelectedIndex = 0;
            }

            serialPort1.DataReceived += SerialPort1_DataReceived;
            //btnRef.Enabled = false;

            foreach (Series serie in chart1.Series.ToList())
            {
                serie.Enabled = false;
            }
        }

        private void SerialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    line = serialPort1.ReadLine();
                    if (readEnabled)
                    {
                        UpdateText(line);
                        UpdateChart(line);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void UpdateChart(string value)
        {
            if (chart1.InvokeRequired)
            {
                chart1.Invoke(new Action(() =>
                {
                    UpdateChart(value);
                }
                ));
            }
            else
            {
                string[] data = value.Split(';');
                if (data.Length == 6)
                {
                    // mean 1
                    mMean1 = 0;
                    mean1.Add(double.Parse(data[0]));                                 // Add new value to mean List.
                    if (mean1.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean1.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean1.Count; i++)            // Loop over all the values of the mean list.
                        mMean1 += mean1[i];                           // Sum up all the values son list.
                    mMean1 = mMean1 / mean1.Count;                     // Divide mMean over the number of elements of List to obtain the mean.
                    // mean 2
                    mMean2 = 0;
                    mean2.Add(double.Parse(data[1]));                                 // Add new value to mean List.
                    if (mean2.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean2.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean2.Count; i++)            // Loop over all the values of the mean list.
                        mMean2 += mean2[i];                           // Sum up all the values son list.
                    mMean2 = mMean2 / mean2.Count;                     // Divide mMean over the number of elements of List to obtain the mean.
                    // mean 3
                    mMean3 = 0;
                    mean3.Add(double.Parse(data[2]));                                 // Add new value to mean List.
                    if (mean3.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean3.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean3.Count; i++)            // Loop over all the values of the mean list.
                        mMean3 += mean3[i];                           // Sum up all the values son list.
                    mMean3 = mMean3 / mean3.Count;                     // Divide mMean over the number of elements of List to obtain the mean.
                    //mean 4
                    mMean4 = 0;
                    mean4.Add(double.Parse(data[3]));                                 // Add new value to mean List.
                    if (mean4.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean4.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean4.Count; i++)            // Loop over all the values of the mean list.
                        mMean4 += mean4[i];                           // Sum up all the values son list.
                    mMean4 = mMean4 / mean4.Count;                     // Divide mMean over the number of elements of List to obtain the mean.
                    //mean 5
                    mMean5 = 0;
                    mean5.Add(double.Parse(data[4]));                                 // Add new value to mean List.
                    if (mean5.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean5.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean5.Count; i++)            // Loop over all the values of the mean list.
                        mMean5 += mean5[i];                           // Sum up all the values son list.
                    mMean5 = mMean5 / mean5.Count;                     // Divide mMean over the number of elements of List to obtain the mean.
                    //mean 6
                    mMean6 = 0;
                    mean6.Add(double.Parse(data[5]));                                 // Add new value to mean List.
                    if (mean6.Count > lMean)                            // Check is mean list has more than 10 elements.
                        mean6.RemoveAt(0);                           // Remove extra value to always have max 10 values on list.
                    for (int i = 0; i < mean6.Count; i++)            // Loop over all the values of the mean list.
                        mMean6 += mean6[i];                           // Sum up all the values son list.
                    mMean6 = mMean6 / mean6.Count;                     // Divide mMean over the number of elements of List to obtain the mean.

                    chart1.Series[0].Points.AddY(mMean1 - reference[0]);
                    chart1.Series[1].Points.AddY(mMean2 - reference[1]);
                    chart1.Series[2].Points.AddY(mMean3 - reference[2]);
                    chart1.Series[3].Points.AddY(mMean4 - reference[3]);
                    chart1.Series[4].Points.AddY(mMean5 - reference[4]);
                    chart1.Series[5].Points.AddY(mMean6 - reference[5]);

                    if (chart1.Series[0].Points.Count >= 100)
                    {
                        chart1.Series[0].Points.RemoveAt(0);
                        chart1.Series[1].Points.RemoveAt(0);
                        chart1.Series[2].Points.RemoveAt(0);
                        chart1.Series[3].Points.RemoveAt(0);
                        chart1.Series[4].Points.RemoveAt(0);
                        chart1.Series[5].Points.RemoveAt(0);
                    }
                    chart1.ChartAreas[0].RecalculateAxesScale();
                }
            }
        }

        private void UpdateText(string value)
        {
            if (textBox1.InvokeRequired)
            {
                textBox1.Invoke(new Action(() =>
                {
                    UpdateText(value);
                }
                ));
            }
            else
            {
                textBox1.AppendText(value);
            }
        }

        private void refreshPorts()
        {
            String[] puertos = SerialPort.GetPortNames();
            cbxPorts.Items.Clear();
            for (int i = 0; i < puertos.Length; i++)
            {
                cbxPorts.Items.Add(puertos[i]);
                cbxPorts.SelectedIndex = 0;
            }

            lblStatus.Text = "Ports Found = " + puertos.Length;
        }


        private void btnRefresh_Click(object sender, EventArgs e)
        {
            refreshPorts();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (btnConnect.Text == "Connect")
            {
                serialPort1.PortName = cbxPorts.SelectedItem.ToString();

                serialPort1.Open();

                serialPort1.ReadExisting();

                lblStatus.Text = "Connection succesfull at port " + serialPort1.PortName;

                btnConnect.Text = "Disconnect";

                btnSndLed.Enabled = true;
                btnSndSettings.Enabled = true;
                btnSoftReset.Enabled = true;

                btnSave.Enabled = false;
                btnStart.Enabled = true;
            }
            else
            {
                timer1.Stop();

                if (serialPort1.IsOpen)
                    serialPort1.Close();

                lblStatus.Text = "Disconnection succesfull at port " + serialPort1.PortName;

                btnConnect.Text = "Connect";
                btnStart.Text = "Start";

                btnStart.Enabled = false;
                btnSndLed.Enabled = false;
                btnSndSettings.Enabled = false;
                btnSoftReset.Enabled = false;
                btnSave.Enabled = true;

            }
        }

        private void btnSndSettings_Click(object sender, EventArgs e)
        {
            byte[] buff = new byte[2];

            if (serialPort1.IsOpen)
            {
                switch (cbxMMode.SelectedIndex)
                {
                    case 0: buff[0] = 0x00; serialPort1.Write(buff, 0, 1); break;
                    case 1: buff[0] = 0x01; serialPort1.Write(buff, 0, 1); break;
                    case 2: buff[0] = 0x02; serialPort1.Write(buff, 0, 1); break;
                    case 3: buff[0] = 0x03; serialPort1.Write(buff, 0, 1); break;
                    default: break;
                }

                if (cbxLIE.SelectedIndex == 0)
                {
                    buff[0] = 0x10;
                    serialPort1.Write(buff, 0, 1);
                }
                else
                {
                    buff[0] = 0x11;
                    serialPort1.Write(buff, 0, 1);
                }

                switch (cbxLIC.SelectedIndex)
                {
                    case 0: buff[0] = 0x12; serialPort1.Write(buff, 0, 1); break;
                    case 1: buff[0] = 0x13; serialPort1.Write(buff, 0, 1); break;
                    case 2: buff[0] = 0x14; serialPort1.Write(buff, 0, 1); break;
                    case 3: buff[0] = 0x15; serialPort1.Write(buff, 0, 1); break;
                    default: break;
                }

                if (cbxLBE.SelectedIndex == 0)
                {
                    buff[0] = 0x20;
                    serialPort1.Write(buff, 0, 1);
                }
                else
                {
                    buff[0] = 0x21;
                    serialPort1.Write(buff, 0, 1);
                }

                switch (cbxLBC.SelectedIndex)
                {
                    case 0: buff[0] = 0x22; serialPort1.Write(buff, 0, 1); break;
                    case 1: buff[0] = 0x23; serialPort1.Write(buff, 0, 1); break;
                    case 2: buff[0] = 0x24; serialPort1.Write(buff, 0, 1); break;
                    case 3: buff[0] = 0x25; serialPort1.Write(buff, 0, 1); break;
                    default: break;
                }
                // Gain
                switch (cbxGain.SelectedIndex)
                {
                    case 0: buff[0] = 0x40; serialPort1.Write(buff, 0, 1); break;
                    case 1: buff[0] = 0x41; serialPort1.Write(buff, 0, 1); break;
                    case 2: buff[0] = 0x42; serialPort1.Write(buff, 0, 1); break;
                    case 3: buff[0] = 0x43; serialPort1.Write(buff, 0, 1); break;
                    default: break;
                }
                // Integration time
                buff[0] = 0x70;
                buff[1] = (byte)numericUpDown1.Value;
                serialPort1.Write(buff, 0, 2);
            }
        }

        private void btnSoftReset_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                byte[] buff = new byte[2];

                buff[0] = 0x30;
                serialPort1.Write(buff, 0, 1);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                byte[] buff = new byte[2];

                buff[0] = 0x60;
                serialPort1.Write(buff, 0, 1);
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (btnStart.Text == "Start")
            {
                timer1.Start();
                btnStart.Text = "Stop";
                readEnabled = true;

                btnRef.Enabled = true;

                foreach (Control item in this.Controls)
                {
                    if (item is CheckBox)
                        ((CheckBox)item).Enabled = false;
                }

                if (cbxCh1.Checked) { chart1.Series[0].Enabled = true; }
                if (cbxCh2.Checked) { chart1.Series[1].Enabled = true; }
                if (cbxCh3.Checked) { chart1.Series[2].Enabled = true; }
                if (cbxCh4.Checked) { chart1.Series[3].Enabled = true; }
                if (cbxCh5.Checked) { chart1.Series[4].Enabled = true; }
                if (cbxCh6.Checked) { chart1.Series[5].Enabled = true; }
            }
            else
            {
                timer1.Stop();
                btnStart.Text = "Start";
                readEnabled = false;

                btnRef.Enabled = false;

                foreach (Control item in this.Controls)
                {
                    if (item is CheckBox)
                        ((CheckBox)item).Enabled = true;
                }
            }
        }

        private void btnRef_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < reference.Length; i++)
                reference[i] = 0;

            readEnabled = false;

            for (int i = 0; i < chart1.Series.Count; i++)
                chart1.Series[i].Points.Clear();

            readEnabled = true;

            MessageBox.Show(this, "Measuring reference, pease wait", "Reference...", MessageBoxButtons.OK, MessageBoxIcon.Information);


            if (chart1.Series[0].Points.Count == 100)
                readEnabled = false;


            for (int i = 0; i < chart1.Series.Count; i++)
            {
                for (int j = 0; j < chart1.Series[i].Points.Count; j++)
                {
                    reference[i] += chart1.Series[i].Points[j].XValue;
                }
                reference[i] /= chart1.Series[i].Points.Count;
            }

            for (int i = 0; i < chart1.Series.Count; i++)
                chart1.Series[i].Points.Clear();


            textBox2.Text = "" + reference[0];

            MessageBox.Show(this, "Reference measured, you can start to work.", "Reference.", MessageBoxButtons.OK, MessageBoxIcon.None);

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();
            chart1.Series[3].Points.Clear();
            chart1.Series[4].Points.Clear();
            chart1.Series[5].Points.Clear();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (chart1.Series[0].Points.Count > 0)
            {
                if (saveFileDialog1.ShowDialog(this) == DialogResult.OK)
                {
                    if (saveFileDialog1.FileName != "")
                    {
                        sw = new StreamWriter(saveFileDialog1.FileName);

                        sw.WriteLine(DateTime.Now.ToLongDateString());
                        sw.WriteLine(DateTime.Now.ToLongTimeString());

                        for (int i = 0; i < chart1.Series[0].Points.Count; i++)
                        {
                            sw.WriteLine(chart1.Series[0].Points[i].YValues[0] + "\t" +
                                         chart1.Series[1].Points[i].YValues[0] + "\t" +
                                         chart1.Series[2].Points[i].YValues[0] + "\t" +
                                         chart1.Series[3].Points[i].YValues[0] + "\t" +
                                         chart1.Series[4].Points[i].YValues[0] + "\t" +
                                         chart1.Series[5].Points[i].YValues[0] + "\t");
                        }

                        sw.Flush();
                        sw.Close();
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Tiff Image|*.tiff";
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    if (sfd.FileName != "")
                    {
                        this.chart1.SaveImage(sfd.FileName, ChartImageFormat.Tiff);
                    }
                }
            }
        }

        private void btnSndLed_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                byte[] buff = new byte[2];

                if (checkRed.Checked) { buff[0] = 0xC0; serialPort1.Write(buff, 0, 1); }
                else { buff[0] = 0xC1; serialPort1.Write(buff, 0, 1); }

                if (checkGreen.Checked) { buff[0] = 0xC2; serialPort1.Write(buff, 0, 1); }
                else { buff[0] = 0xC3; serialPort1.Write(buff, 0, 1); }

                if (checkBlue.Checked) { buff[0] = 0xC4; serialPort1.Write(buff, 0, 1); }
                else { buff[0] = 0xC5; serialPort1.Write(buff, 0, 1); }

            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
