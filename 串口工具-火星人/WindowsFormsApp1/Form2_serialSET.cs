using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static WindowsFormsApp1.Class1_json;
using System.IO.Ports;

namespace WindowsFormsApp1
{
    /// <summary>
    /// 串口设置
    /// </summary>
    public partial class Form2_serialSET : Form
    {
        public Form2_serialSET()
        {
            InitializeComponent();
        }
        //实例化一个获取json配置的对象
        My_config my_Config = ReadFile();

        /// <summary>
        /// 从json配置文件中导入上次默认配置
        /// </summary>
        /// <param name="my_Config">获取json配置的对象</param>
        private void Json_load(My_config my_Config)
        {
            //从json配置文件中导入上次默认配置
            comboBox1.Text = my_Config.Serial_config.intBaudRate;
            comboBox2.Text = my_Config.Serial_config.intDataBits + " bit";
            comboBox3.Text = my_Config.Serial_config.intStopBits + " bit";
            comboBox4.Text = my_Config.Serial_config.intParity;
        }

        /// <summary>
        /// 把配置写入json文件的方法
        /// </summary>
        private void Json_write()
        {
            var config = new My_config();
            var serial_config = new Serial_config();
            

            serial_config = new Serial_config()
            {
                intBaudRate = comboBox1.Text,
                intDataBits = comboBox2.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0],
                intStopBits = comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0],
                intParity = comboBox4.Text
            };
            config = new My_config
            {
                Command_config = my_Config.Command_config,
                Serial_config = serial_config,
                Command_Name = my_Config.Command_Name,
                List_Name = my_Config.List_Name,
                List_Value = my_Config.List_Value,
                CheckCode_value = my_Config.CheckCode_value
            };
            WriteFile(config);
        }

        /// <summary>
        /// 把修改的串口配置推送给串口实例的属性
        /// </summary>
        private void Serial_put()
        {
            Form1.intBaudRate = int.Parse(comboBox1.Text);
            Form1.intDataBits = int.Parse(comboBox2.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0]);
            if (comboBox3.Text == "1.5 bit")
            {
                Form1.intStopBits = StopBits.OnePointFive;
            }
            else
            {
                Form1.intStopBits = (StopBits)int.Parse(comboBox3.Text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0]);
            }
            if (comboBox4.Text == "none")
                Form1.intParity = Parity.None;
            if (comboBox4.Text == "odd")
                Form1.intParity = Parity.Odd;
            if (comboBox4.Text == "even")
                Form1.intParity = Parity.Even;
            if (comboBox4.Text == "mark")
                Form1.intParity = Parity.Mark;
            if (comboBox4.Text == "space")
                Form1.intParity = Parity.Space;
        }

        /// <summary>
        /// 运行后自动加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_serialSET_Load(object sender, EventArgs e)
        {
            this.Text = "火星人：" + "串口设置";
            //比特率
            comboBox1.Items.Add(110);
            comboBox1.Items.Add(300);
            comboBox1.Items.Add(600);
            comboBox1.Items.Add(1200);
            comboBox1.Items.Add(2400);
            comboBox1.Items.Add(4800);
            comboBox1.Items.Add(9600);
            comboBox1.Items.Add(14400);
            comboBox1.Items.Add(19200);
            comboBox1.Items.Add(38400);
            comboBox1.Items.Add(57600);
            comboBox1.Items.Add(115200);
            comboBox1.Items.Add(230400);
            comboBox1.Items.Add(460800);
            comboBox1.Items.Add(921600);
            //数据位
            comboBox2.Items.Add("5 bit");
            comboBox2.Items.Add("6 bit");
            comboBox2.Items.Add("7 bit");
            comboBox2.Items.Add("8 bit");
            //停止位
            comboBox3.Items.Add("0 bit");
            comboBox3.Items.Add("1 bit");
            comboBox3.Items.Add("1.5 bit");
            comboBox3.Items.Add("2 bit");
            //奇偶校验位
            comboBox4.Items.Add("none");
            comboBox4.Items.Add("odd");
            comboBox4.Items.Add("even");
            comboBox4.Items.Add("mark");
            comboBox4.Items.Add("space");
            //默认值设置
            /*comboBox1.SelectedIndex = 6;    //比特率默认9600
            comboBox2.SelectedIndex = 3;    //数据位默认8bit
            comboBox3.SelectedIndex = 1;    //停止位默认1bit
            comboBox4.SelectedIndex = 0;    //奇偶校验位默认：无校验*/

            //从json配置文件my_Config中导入上次默认配置
            Json_load(my_Config);
        }

        /// <summary>
        /// 确认按键：点击后修改当前串口配置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //点击确认按键后把当前串口配置写入配置文件Config_json.json中
            Json_write();
            //点击确认按键后把当前串口配置传入当前打开的串口实例中
            Serial_put();
            this.DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// 取消按钮：点击取消新修改的串口配置，保留之前配置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
