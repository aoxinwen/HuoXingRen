using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;

namespace WindowsFormsApp1
{
    /// <summary>
    /// 新建连接
    /// </summary>
    public partial class Form2_newConnect : Form
    {
        public Form2_newConnect()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 开启“选择串口”窗口的时候自动加载当前串口号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_newConnect_Load(object sender, EventArgs e)
        {
            this.Text = "火星人：" + "新建连接";
            //获取当前连接的串口
            string[] ports = SerialPort.GetPortNames();
            foreach(string port in ports)
            {
                this.comboBox1.Items.Add(port);
            }
           /* if (Form1.sp.IsOpen)
            {
                this.comboBox1.Items.Remove(Form1.sp.PortName);
            }*/
            //默认选择第一个串口
            this.comboBox1.SelectedIndex = 0;
        }


        /// <summary>
        /// “取消”按钮：点击取消取消弹窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        /// <summary>
        /// “确认”按键：点击处理数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Form1.strPortName = this.comboBox1.Text;
            this.DialogResult = DialogResult.OK;
        }
    }
}
