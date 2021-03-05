using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using static WindowsFormsApp1.Class1_json;
using System.Diagnostics;
using System.Threading;
using static WindowsFormsApp1.Class_excel;


// 主窗体
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string strPortName = "";                // 初始化串口名称
        public static int intBaudRate = 9600;               // 初始化串口波特率
        public static Parity intParity = Parity.None;         // 初始化串口校验位（奇偶校验 0 1 2 3 4）
        public static int intDataBits = 8;                    // 初始化串口数据位 (8)
        public static StopBits intStopBits = StopBits.One;    // 初始化串口停止位 (0 1 2 3)
        public static string strFrameHead = "";               // 初始化数据帧头
        public static string strFrameTail = "";               // 初始化数据帧尾
        public static string strFrameCheckCode = "";          // 初始化数据校验模式（CRC校验）
        public static int intFrameSequential;                 // 初始化数据包 时序号 在列表的索引
        //My_config my_Config = ReadFile();                     // 初始化json数据
        public static List<string> listview_name = new List<string>();           //listview里面显示的表头
        public static List<string> listview_value = new List<string>();          //listview里面显示的内容
        public static List<string> listview_check_value = new List<string>();    //listview里面内容需要校验的项
        public static List<string> listcombox2_value = new List<string>(new string[10]); //待添加的选项内容


        /// <summary>
        /// 打开串口
        /// </summary>
        private void Open_serial()
        {
            try
            {
                //串口属性赋值：串口名、波特率、数据位、停止位、奇偶校验等
                this.serialPort1.PortName = strPortName;
                this.serialPort1.BaudRate = intBaudRate;
                this.serialPort1.DataBits = intDataBits;
                this.serialPort1.StopBits = intStopBits;
                this.serialPort1.Parity = intParity;
                //this.serialPort1.ReadTimeout = 1000;
                //如果串口已经打开，弹窗提示已经打开
                if (this.serialPort1.IsOpen)
                {
                    MessageBox.Show(strPortName + "已经打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //否则直接打开
                else
                    this.serialPort1.Open(); 
                //窗体名称修改为：已经打开的串口+名字
                this.Text = strPortName + "-" + "串口小工具-火星人";
                //“新建连接”菜单置灰，激活“断开连接”菜单
                this.新建连接ToolStripMenuItem.Enabled = false;
                this.断开连接ToolStripMenuItem.Enabled = true;
            }
            catch
            {
                MessageBox.Show("打开失败", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 关闭已经打开的串口
        /// </summary>
        private void close_serial()
        {
            try
            {
                //如果打开状态则直接关闭
                if (this.serialPort1.IsOpen)
                {
                    this.serialPort1.Close();
                }
                //否则提示已经关闭
                else
                {
                    MessageBox.Show(strPortName + "已经关闭", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //窗体名称修改：默认名称
                this.Text = "串口小工具-[未连接]火星人";
                //激活“新建连接”菜单，置灰“断开连接”菜单
                this.新建连接ToolStripMenuItem.Enabled = true;
                this.断开连接ToolStripMenuItem.Enabled = false;
            }
            catch
            {
                MessageBox.Show("断开连接失败", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 读取串口消息并且显示在richbox上
        /// </summary>
        private StringBuilder Read_display_data()
        {
            string data;
            StringBuilder sb = new StringBuilder();
            //串口可读取的字节数
            int n = this.serialPort1.BytesToRead;
            byte[] buf = new byte[n];
            this.serialPort1.Read(buf, 0, n);
            //如果选择HEX接收
            if (this.hEX接收ToolStripMenuItem.Checked)
            {
                for (int i = 0; i < n; i++)
                {
                    if (buf[i] < 16)
                        data = "0" + Convert.ToString(buf[i], 16).ToUpper() + " ";
                    else
                        data = Convert.ToString(buf[i], 16).ToUpper() + " ";
                    sb.Append(data);
                }
            }
            //如果选择ASCII newline接收
            else if (this.aSCIIToolStripMenuItem2.Checked)
            {
                data = this.serialPort1.ReadLine();
                sb.Append(data);
            }
            //选择ASCII接收
            else
            {
                data = Encoding.Default.GetString(buf);
                sb.Append(data);
            }
            return sb;
        }

        /// <summary>
        /// 向串口写入数据data
        /// </summary>
        /// <param name="data"></param>
        private bool Write_data(string data)
        {
            string hexstring;
            bool result = false;
            if (this.rn发送ToolStripMenuItem.Checked)
                hexstring = data.Replace(" ", "") + "\r\n";//去掉字符串内的空格+上换行符
            else
                hexstring = data.Replace(" ", "");//去掉字符串内的空格
            byte[] buff = new byte[hexstring.Length/2];//每两位转换一次
            //如果串口打开
            if (this.serialPort1.IsOpen)
            {
                //如果选择ASCII发送
                if (this.aSCIIToolStripMenuItem1.Checked)
                {
                    buff = Encoding.ASCII.GetBytes(hexstring);
                }                   
                //如果选择HEX发送
                else
                {
                    try
                    {
                        for (int i = 0; i < buff.Length; i++)
                        {
                            //把数字字符串转成16进制byte
                            buff[i] = Convert.ToByte(hexstring.Substring(i * 2, 2), 16);                            
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "转换成HEX错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                try
                {
                    this.serialPort1.Write(buff, 0, buff.Length);
                    result = true;
                }
                catch (Exception e)
                {                   
                    MessageBox.Show(e.Message, "写数据失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("串口未打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return result;
        }

        /// <summary>
        /// 把list非null元素挨个添加进box的items
        /// </summary>
        /// <param name="box">要添加的box</param>
        /// <param name="list">被添加的list</param>
        private static void Combox_add(ComboBox box, List<string> list)
        {
            box.Items.Clear();
            foreach (string li in list)
            {
                if (li != null)
                {
                    box.Items.Add(li);
                }
            }
            if (box.Items.Count > 0)
            {
                box.SelectedIndex = 0;
            }
            else
                box.Text = null;
        }

        /// <summary>
        /// 初始化json值
        /// </summary>
        private void Json_load()
        {
            My_config my_Config = ReadFile();
            //List_Name
            listview_name = new List<string>(new string[] { my_Config.List_Name.one, my_Config.List_Name.two, my_Config.List_Name.three, my_Config.List_Name.four, my_Config.List_Name.five, my_Config.List_Name.six, my_Config.List_Name.seven, my_Config.List_Name.eight, my_Config.List_Name.nine, my_Config.List_Name.ten });

            //List_Value
            listview_value = new List<string>(new string[] { my_Config.List_Value.one, my_Config.List_Value.two, my_Config.List_Value.three, my_Config.List_Value.four, my_Config.List_Value.five, my_Config.List_Value.six, my_Config.List_Value.seven, my_Config.List_Value.eight, my_Config.List_Value.nine, my_Config.List_Value.ten });

            //CheckCode_value
            listview_check_value = new List<string>(new string[] { my_Config.CheckCode_value.one, my_Config.CheckCode_value.two, my_Config.CheckCode_value.three, my_Config.CheckCode_value.four, my_Config.CheckCode_value.five, my_Config.CheckCode_value.six, my_Config.CheckCode_value.seven, my_Config.CheckCode_value.eight, my_Config.CheckCode_value.nine, my_Config.CheckCode_value.ten });

            //先清除comboBox2（需要用户自己添加的项）的内容
            this.comboBox2.Items.Clear();
            //如果选项名存在，而选项值为“”或者“请输入..”表示该项为用户自己输入
            for (int i=0;i< listview_name.Count;i++)
            {               
                if (listview_value[i] == "" || listview_value[i] == "请输入..")
                {
                    listcombox2_value[i] = listview_name[i] + ": ";
                }
                else
                    listcombox2_value[i] = null;

                if (listview_name[i] == "帧头")
                {
                    strFrameHead = string.Empty;
                    for (int j = 0; j < listview_value[i].Length / 2; j++)
                    {
                        strFrameHead += listview_value[i].Substring(j * 2, 2) + " ";
                    }
                }
                if (listview_name[i] == "帧尾")
                {
                    strFrameTail = listview_value[i];

                    strFrameTail = string.Empty;

                    for (int k = 0; k < listview_value[i].Length/2; k++)
                    {
                        strFrameTail += listview_value[i].Substring(k * 2, 2) + " ";
                    }
                }
            }
            //需要手动添加的项值都添加到comboBox2内，并默认显示第一项
            Combox_add(this.comboBox2, listcombox2_value);           
        }

        /// <summary>
        /// 把data按照mode的模式计算出校验码
        /// </summary>
        /// <param name="mode">模式</param>
        /// <param name="data">数据</param>
        /// <returns>返回校验码</returns>
        private static string pyCheck(string mode, string data)
        {
            string pyexePath = @".\crcApi.exe";
            Process p = new Process();
            p.StartInfo.FileName = pyexePath;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.Arguments = mode + " " + data;
            p.Start();
            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();
            // Console.Write(output);
            p.Close();            
            switch (output.Replace("\r\n", "").Length)
            {
                case 0:
                    output = "0000";
                    break;
                case 1:
                    output = "000" + output;
                    break;
                case 2:
                    output = "00" + output;
                    break;
                case 3:
                    output = "0" + output;
                    break;
            }            
            return output;
        }

        /// <summary>
        /// 根据bit的位数自动生成一个随机数（16进制显示）
        /// </summary>
        /// <param name="bit"></param>
        /// <returns></returns>
        private static string Order_number(string bit)
        {
            string order_number;
            int i;
            Random r = new Random();
            if (bit == "0")
            {
                order_number = "";
            }
            else if (bit == "1")
            {
                i = r.Next(0, 256);
                if (i < 16)
                    order_number = "0" + Convert.ToString(i, 16).ToUpper();
                else
                    order_number = Convert.ToString(i, 16).ToUpper();
            }
            else if (bit == "2")
            {
                i = r.Next(0, 65536);
                if (i < 16)
                    order_number = "000" + Convert.ToString(i, 16).ToUpper();
                else if (i < 256)
                    order_number = "00 " + Convert.ToString(i, 16).ToUpper();
                else if (i < 4096)
                    order_number = "0" + Convert.ToString(i, 16).ToUpper();
                else
                    order_number = Convert.ToString(i, 16).ToUpper();
            }
            else if (bit == "3")
            {
                i = r.Next(0, 16777216);
                if (i < 16)
                    order_number = "00000" + Convert.ToString(i, 16).ToUpper();
                else if (i < 256)
                    order_number = "0000 " + Convert.ToString(i, 16).ToUpper();
                else if (i < 4096)
                    order_number = "000" + Convert.ToString(i, 16).ToUpper();
                else if (i < 65536)
                    order_number = "00" + Convert.ToString(i, 16).ToUpper();
                else if (i < 1048576)
                    order_number = "0" + Convert.ToString(i, 16).ToUpper();
                else
                    order_number = Convert.ToString(i, 16).ToUpper();
            }
            else
                order_number = bit;
            return order_number;
        }

         /// <summary>
        /// 根据box内容返回计算出的整个报文
        /// </summary>
        /// <param name="box"></param>
        /// <returns></returns>
        private static string Out_transform(List<string> box)
        {
            //初始化转换输出的字符串
            string outdata = "";
            //初始化时序号
            string sq_number = null;
            //初始化校验码
            string pydata = null;
            //com_index:第几个待输出的项
            int com_index = 0;
            int com_index1 = 0;
            for (int i = 0; i < listview_name.Count; i++)
            {
                if (canstop_create)
                    break;
                if (listview_value[i] != null && listview_name[i] != "校验" && listview_name[i] != "时序")
                {
                    if (listview_value[i] != "" && listview_value[i] != "请输入..")
                        outdata += listview_value[i];
                    else
                    {
                        //outdata += box.Items[com_index].ToString().Split(' ')[1];
                        outdata += box[com_index].Split(' ')[1];
                        com_index += 1;
                    }
                }
                else if (listview_value[i] != null && listview_name[i] == "时序")
                {
                    if (sq_number == null)
                        sq_number = Order_number(listview_value[i]);
                    outdata += sq_number;
                }
                else if (listview_value[i] != null && listview_name[i] == "校验")
                {                    
                    for (int j = 0; j < listview_check_value.Count; j++)
                    {
                        if (canstop_create)
                            break;
                        if (listview_check_value[j] != null && listview_name[j] == "时序")
                        {
                            if (sq_number == null)
                                sq_number = Order_number(listview_value[j]);
                            pydata += sq_number;
                        }
                        else if (listview_check_value[j] == "" || listview_check_value[j] == "请输入..")
                        {
                            //pydata += box.Items[com_index1].ToString().Split(' ')[1];
                            pydata += box[com_index1].Split(' ')[1];
                            com_index1 += 1;
                        }
                        else if (listview_check_value[j] != null)
                            pydata += listview_check_value[j];
                    }
                    pydata = pyCheck(listview_value[i], pydata);
                    outdata += pydata.Replace("\r\n","").ToUpper();
                }
            }
            return outdata.Replace("\r\n", "").ToUpper();
        }

        /// <summary>
        /// 把list的非null元素挨个添加进listview的Columns
        /// </summary>
        /// <param name="lv">要添加的ListView</param>
        /// <param name="list">被添加的list</param>
        private static void Listview_Columns_add(ListView lv, List<string> list)
        {
            lv.Clear();
            lv.BeginUpdate();
            foreach (string li in list)
            {
                if (li != null)
                {
                    lv.Columns.Add(li,120);
                }
            }
            lv.EndUpdate();
        }

        /// <summary>
        /// 把list的非null元素挨个添加进listview的Items
        /// </summary>
        /// <param name="lv">要添加的ListView</param>
        /// <param name="list">被添加的list</param>
        private static void Listview_value_add(ListView lv, List<string> list)
        {
            ListViewItem lvi = new ListViewItem();
            lvi.UseItemStyleForSubItems = false;  //设置单元格可以修改颜色和格式
            lv.BeginUpdate();
            lvi.ImageIndex = 0;
            lvi.Text = test_time.ToString();
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] != null)
                {
                    lvi.SubItems.Add(list[i]);
                }
            }
            if (list.Contains("fail"))
            {
                lvi.BackColor = Color.Red;
                for (int i =0; i<lvi.SubItems.Count;i++)
                {
                    lvi.SubItems[i].BackColor = Color.LightPink;
                }
            }
            if (list.Contains("pass"))
            {
                lvi.BackColor = Color.Green;
                for (int i = 0; i < lvi.SubItems.Count; i++)
                {
                    lvi.SubItems[i].BackColor = Color.LightGreen;
                }
            }
            lv.Items.Add(lvi);
            lv.EndUpdate();
            lv.Items[lv.Items.Count - 1].EnsureVisible();
        }

        /// <summary>
        /// 开启主窗体后自动加载的内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            //默认“断开连接”菜单置灰
            this.断开连接ToolStripMenuItem.Enabled = false;
            //实例化 串口选择 窗体对象：默认首先打开 串口选择 窗体
            Form2_newConnect dlg = new Form2_newConnect();
            //如果 串口选择 窗体关闭返回值是：ok 就打开串口
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Open_serial();
            }           
            this.splitContainer2.Panel2Collapsed = true;//命令窗口区域默认隐藏*/
            this.splitContainer1.Panel2Collapsed = true;//测试打印区域默认隐藏           
            this.hEX接收ToolStripMenuItem.Checked = true;    //默认ASCII接收
            this.hEX发送ToolStripMenuItem.Checked = true;   //默认ASCII发送
            this.comboBox2.DropDownStyle = ComboBoxStyle.DropDownList; //默认不可输入
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;//默认不可输入
            comboBox1.Items.Add("请添加...");
            comboBox1.SelectedIndex = 0;
            this.listView1.View = View.Details;
            this.listView1.SmallImageList = this.imageList1;
            this.textBox3.Text = "1000";   //默认定时1000ms发送
            this.textBox5.Text = "1";      //默认循环次数1
            this.textBox6.Text = "1000";   //默认间隔时间1000ms
            this.textBox4.ReadOnly = true;
            this.richTextBox2.ReadOnly = true;//显示log打印只读
            Json_load();          
        }

        /// <summary>
        /// 点击菜单：文件-新建连接：新建一个串口连接
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 新建连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2_newConnect dlg = new Form2_newConnect();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                //打开串口
                Open_serial();
            }           
        }

        /// <summary>
        /// 点击菜单：文件-断开连接：断开串口连接
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 断开连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //关闭串口
            close_serial();
        }

        /// <summary>
        /// 点击菜单：文件-退出：退出程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //退出程序
            this.Close();
        }

        /// <summary>
        /// 清空屏幕显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 清空屏幕ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.richTextBox2.Clear();
        }

        /// <summary>
        /// 串口配置的设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 串口设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2_serialSET dlg = new Form2_serialSET();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (this.serialPort1.IsOpen)
                {
                    this.serialPort1.Close();
                    Open_serial();
                }
            }
        }

        /// <summary>
        /// 报文设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 报文设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2_messageSET dlg = new Form2_messageSET();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Json_load();                
            }
        }

        private string com_content; //初始化接收内容的集合       
        
        /// <summary>
        /// 打印log的线程
        /// </summary>
        private void thread_log()
        {
            Thread.Sleep(2000);//等待2s时间等接收完成
            while (!canstop_log)
            {                
                if (com_content == null)
                {
                    continue;
                }                
                //根据报文 报头 分割收到的数据存在数组receive_arr内
                string[] receive_arr = com_content.Split(new[] { strFrameHead.ToUpper() }, StringSplitOptions.RemoveEmptyEntries);
                List<string> receive_list = new List<string>();//初始化一个列表存放完整的接收报文                
                foreach (string tent in receive_arr)
                {                    
                    if (tent.EndsWith(strFrameTail.ToUpper()))
                    {
                        receive_list.Add(strFrameHead.ToUpper() + tent);                                          
                    }                   
                }

                //遍历所有excel（发送数据）的发送标志位（已经发送为1）
                for (int i = 1; i < case_content.GetLength(0); i++)
                {
                    //Console.WriteLine("开始i:" + i);
                    if (canstop_log)
                        break;
                    if (Convert.ToInt32(case_content[i, case_content.GetLength(1) - 2])  >= test_time && Convert.ToInt32(case_content[i, case_content.GetLength(1) - 1]) < test_time) //发送标志位是1的情况
                    {
                        case_content[i, case_content.GetLength(1) - 1] = test_time.ToString();
                        List<string> _list = new List<string>(new string[7]);//显示在listview的内容
                                                                             //显示前5个数据                       
                        for (int j = 0; j < 5; j++)
                        {
                            _list[j] = case_content[i, j];
                        }
                        if (receive_list != null)
                        {
                            _list[5] = "没收到";
                            _list[6] = "fail";
                            foreach (string re_tent in receive_list)
                            {
                                if (re_tent != null)
                                {                                    
                                    //如果时序号相同
                                    if (re_tent.Replace(" ", "").IndexOf(case_content[i, 5]) == Convert.ToInt32(case_content[i, 6]))
                                    {                                        
                                        _list[5] = re_tent.Replace(" ", "");
                                        if (re_tent.Replace(" ", "") == case_content[i, 4])
                                        {
                                            _list[6] = "pass";
                                            //com_content=com_content.Replace(re_tent, "");
                                            com_content = com_content.Remove(com_content.IndexOf(re_tent), re_tent.Length);                                            
                                            break;
                                        }
                                        else
                                            _list[6] = "fail";
                                    }
                                }
                            }
                        }
                        else
                        {
                            _list[5] = "列表为空";
                            _list[6] = "fail";
                        }
                        if (!canstop_log)
                            this.BeginInvoke(new MethodInvoker(delegate {
                                Listview_value_add(this.listView1, _list);
                            }));
                        Thread.Sleep(time);
                        if (i == case_content.GetLength(0)-1)
                        {
                            if (test_time == num_times)
                            {
                                canstop_log = true;
                                this.BeginInvoke(new MethodInvoker(delegate {
                                    this.button5.Text = "完成测试";
                                }));
                            }
                            else
                                test_time += 1;
                        }                        
                        break;
                    }
                }                
            }           
        }

        /// <summary>
        /// 串口数据接收事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string content =  Read_display_data().ToString();
            //Console.WriteLine("content:" + content);            
            if (this.button5.Text == "开始测试")
            {
                this.BeginInvoke(new MethodInvoker(delegate
                {
                    this.richTextBox2.AppendText(content);
                }));
            }
            else
            {
                com_content += content;               
            }           
        }     

        /// <summary>
        /// 当菜单“命令窗口”选择后弹出输入和测试的命令窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 命令窗口ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.splitContainer2.Panel2Collapsed = !this.splitContainer2.Panel2Collapsed;
        }

        /// <summary>
        /// HEX接收：选中则ASCII接收取消选中，反之亦然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hEX接收ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.hEX接收ToolStripMenuItem.Checked = !this.hEX接收ToolStripMenuItem.Checked;
            this.aSCIIToolStripMenuItem.Checked = !this.hEX接收ToolStripMenuItem.Checked;
            //hex接收选中的时候 newline接收置灰，否则激活
            if (this.hEX接收ToolStripMenuItem.Checked)
                this.aSCIIToolStripMenuItem2.Enabled = false;
            else
                this.aSCIIToolStripMenuItem2.Enabled = true;
        }

        /// <summary>
        /// HEX发送：选中则ASCII发送取消选中，反之亦然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hEX发送ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.hEX发送ToolStripMenuItem.Checked = !this.hEX发送ToolStripMenuItem.Checked;
            this.aSCIIToolStripMenuItem1.Checked = !this.hEX发送ToolStripMenuItem.Checked;
        }

        /// <summary>
        /// ASCII接收：选中则HEX接收取消选中，反之亦然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aSCIIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.aSCIIToolStripMenuItem.Checked = !this.aSCIIToolStripMenuItem.Checked;
            this.hEX接收ToolStripMenuItem.Checked = !this.aSCIIToolStripMenuItem.Checked;
        }

        /// <summary>
        /// ASCII发送：选中则HEX发送取消选中，反之亦然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aSCIIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.aSCIIToolStripMenuItem1.Checked = !this.aSCIIToolStripMenuItem1.Checked;
            this.hEX发送ToolStripMenuItem.Checked = !this.aSCIIToolStripMenuItem1.Checked;
        }

        /// <summary>
        /// 窗体关闭的时候提示，如果确认关闭则关闭已经打开的串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("你确定要关闭吗！", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                if (this.serialPort1.IsOpen)
                {
                    canstop = true;//关闭窗体的话主动关闭 循环发送 的线程
                    canstop_test = true;//关闭窗体的话主动关闭 开始测试 的线程
                    canstop_case = true;//关闭窗体的话主动关闭 生成用例 的线程
                    canstop_log = true;//关闭窗体的话主动关闭 打印log 的线程
                    canstop_create = true;//关闭窗体的话主动关闭 生成报文 的线程
                    this.serialPort1.DiscardInBuffer();
                    this.serialPort1.Close();
                }
                e.Cancel = false;
            }
            else
                e.Cancel = true;
        }

        /// <summary>
        /// 发送按钮：点击发送把richTextBox1内容发送到串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Write_data(this.richTextBox1.Text);
        }

        private List<string> input = new List<string>();
        private static bool canstop_create = false;

        /// <summary>
        /// 新建生成的线程
        /// </summary>
        private void thread_create()
        {          
            string output;
            output = Out_transform(input);           
            this.BeginInvoke(new MethodInvoker(delegate
            {
                this.richTextBox1.Text = output;
            }));
        }

        /// <summary>
        /// 生成按钮：点击把textbox1内容转换成需要的报文，写入到richbox1内等待发送
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {         
            for(int i=0;i<this.comboBox2.Items.Count;i++)
            {
                input.Add(this.comboBox2.Items[i].ToString());
            }
            Thread th1 = new Thread(thread_create);
            canstop_create = false;
            th1.Start();
        }

        /// <summary>
        /// textBox1键盘按键点击事件：点击enter按键后txbox1内容赋值给待输入选项
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //当点击的是enter时候
            if (e.KeyChar == 13)
            {
                if (textBox1.Text != string.Empty && comboBox2.Items.Count > 0)
                    this.comboBox2.Items[comboBox2.SelectedIndex] = this.comboBox2.Items[comboBox2.SelectedIndex].ToString().Split(' ')[0]+" " + textBox1.Text;
            }
        }

        private static System.Timers.Timer t;   //初始化一个基于服务器的定时器t：定时发送  按钮
        private string text;  //初始化要左边发送的内容

        /// <summary>
        /// 判断是否是数字 字符串，是则返回true，否则返回false
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private bool Is_int(string text)
        {
            bool time;
            try
            {
                Convert.ToInt32(text);
                time = true;
            }
            catch
            {
                time = false;
            }
            return time;
        }

        /// <summary>
        /// 创建一个计时器，在一定时间间隔内循环执行thout
        /// </summary>
        private void ThreadMethod()
        {
            int time;
            text = richTextBox1.Text.Replace(" ", "");
            try
            {
                time = Convert.ToInt32(this.textBox3.Text);
                byte[] test = new byte[text.Length/2];
                for (int i=0;i<test.Length;i++)
                {
                    test[i] = Convert.ToByte(text.Substring(i * 2, 2), 16);
                    Console.WriteLine("test[i]:::" + test[i]);
                }                
                t = new System.Timers.Timer(time);
                
                if (this.checkBox1.Checked && this.serialPort1.IsOpen)
                {
                    this.textBox3.ReadOnly = true; //间隔时间运行过程中不可修改
                    t.Elapsed += new System.Timers.ElapsedEventHandler(thout);
                    t.Enabled = true;
                    t.Start();
                }
                else if (this.checkBox1.Checked && !this.serialPort1.IsOpen)
                {
                    this.checkBox1.Checked = false;
                    MessageBox.Show("串口未打开", "提示", MessageBoxButtons.OK);
                }                   
                else
                {
                    this.textBox3.ReadOnly = false;
                    t.Enabled = false;
                    t.Stop();
                }                
            }
            catch
            {
                this.checkBox1.Checked = false;
                MessageBox.Show("输入非数字", "提示", MessageBoxButtons.OK);
            }          
        }

        /// <summary>
        /// 计时器t的执行事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void thout(object source, System.Timers.ElapsedEventArgs e)
        {
            Write_data(text);                        
        }

        /// <summary>
        /// 定时发送选中的时候循环发送
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if(this.checkBox1.Checked)
            {
                this.button2.Enabled = false; //发送按钮暂时失效
                ThreadMethod();
            }
            else
                this.button2.Enabled = true; //发送按钮生效
        }

        /// <summary>
        /// 根据用例计算校验值后补全报文
        /// </summary>
        private static string Get_message(string content)
        {
            //初始化转换输出的字符串
            string outdata = "";
            //初始化校验码
            string pydata = "";
            for (int i=0; i < listview_name.Count; i++)
            {
                if (listview_check_value[i] == null && listview_value[i] != null)
                {
                    if (listview_name[i] == "校验")
                    {
                        pydata = pyCheck(listview_value[i], content);
                        outdata += content;
                        outdata += pydata;
                    }
                    else
                        outdata += listview_value[i];
                }
            }
            return outdata.Replace("\r\n", "").ToUpper();
        }

        /// <summary>
        /// 根据excel内容补全完整报文后重新保存excel
        /// </summary>
        /// <param name="fileName"></param>
        private void Get_excel(string fileName)
        {
            if (Open_excel(fileName)) //打开excel表格
            {
                if (Get_sheet(wb.GetSheetIndex("测试用例")))
                {
                    for (int i=1;i< exsheet.LastRowNum + 1;i++)   //从第2行开始到最后一行，因为第1行是标题
                    {
                        if (!canstop_case) //停止标志位false
                        {
                            if (Get_row(i))  //获取第i+1行
                            {
                                if (Get_cell(i, 1)) //获取第i+1行第2个单元格
                                {
                                    string content = Get_message(cell.StringCellValue);  //第i+1行第2个单元格内容补全报文后赋值给content
                                    if (row.GetCell(2) == null)   //获取第i+1行第3个单元格
                                        row.CreateCell(2).SetCellValue(content);
                                    else
                                        row.GetCell(2).SetCellValue(content);

                                }
                                if (Get_cell(i, 3))
                                {
                                    string content = Get_message(cell.StringCellValue);
                                    if (row.GetCell(4) == null)
                                        row.CreateCell(4).SetCellValue(content);
                                    else
                                        row.GetCell(4).SetCellValue(content);                                                                    
                                }
                                Write_sheet(fileName);
                                decimal a = Math.Round((decimal)i / (exsheet.LastRowNum + 1), 2);
                                this.BeginInvoke(new MethodInvoker(delegate
                                {
                                    this.progressBar1.Value = (int)(a * 100);
                                }));                                
                            }
                        }
                        else
                            break;                        
                    }
                    this.BeginInvoke(new MethodInvoker(delegate
                    {
                        this.button4.Text = "生成用例";
                        this.progressBar1.Value = 100;
                    }));
                }
                Close_excel();
            }             
        }

        private string filepath;//初始化case路径
        private bool canstop_case = false;//生成用例的线程停止标志位
        private bool canstop_test = false;//测试用例的线程停止标志位 
        private bool canstop_log = false;//测试用例的线程停止标志位
        private string[,] case_content; //初始化一个多维数组存excel内容
        private static int test_time;   //初始化case测试的当前次数      

        /// <summary>
        /// 新建线程：生成用例的方法
        /// </summary>
        private void Thread_testcase()
        {
            Get_excel(filepath);
        }

        /// <summary>
        /// 新建线程：开始测试的方法
        /// </summary>
        private void Thread_beginTest()
        {
            if (case_content !=null)
                Array.Clear(case_content, 0, case_content.Length);//开始先清除发送列表

            //挨个把发送项添加进发送列表send_content里                                               
            if (Open_excel(filepath)) //打开excel表格
            {
                if (Get_sheet(wb.GetSheetIndex("测试用例")))  //获取名字叫做“测试用例”的sheet
                {
                    if (Get_row(0))
                    {
                        case_content = new string[exsheet.LastRowNum + 1, row.LastCellNum + 3];//多出来的两列存放发送和接收的标志位
                    }
                    for (int i = 0; i < exsheet.LastRowNum + 1; i++) //从第1行开始到最后一行
                    {
                        if (Get_row(i))  //获取第i+1行,从第1行开始
                        {
                            for (int j = 0; j < case_content.GetLength(1); j++)
                            {                               
                                if (Get_cell(i, j))
                                {
                                    case_content[i, j] = cell.StringCellValue.ToUpper();
                                }                                
                            }
                            case_content[i, case_content.GetLength(1) - 2] = "0";    //初始化发送标志位
                            case_content[i, case_content.GetLength(1) - 1] = "0";    //初始化接收标志位
                        }
                    }
                }
                Close_excel();
            }            
            for (int i = 0; i < num_times; i++)
            {               
                for (int j=1;j< case_content.GetLength(0);j++)
                {
                    if (!canstop_test)
                    {
                        Write_data(case_content[j, 2]);               
                        Thread.Sleep(time);
                        case_content[j, case_content.GetLength(1) - 2] = (i+1).ToString();
                    }
                    else
                        break;
                }               
                if (canstop_test)
                    break;
            }
            //循环完成后托管button7显示发送
            /*if (!canstop_test)
            {
                this.BeginInvoke(new MethodInvoker(delegate
                {
                    this.button5.Text = "开始测试";
                }));
            }*/
        }

        /// <summary>
        /// 开始测试按钮：点击后开始测试
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {      
            
            if (this.textBox4.Text != null && this.textBox4.Text != "")
            {                
                try
                {
                    time = Convert.ToInt32(this.textBox6.Text);
                    num_times = Convert.ToInt32(this.textBox5.Text);
                    filepath = this.textBox4.Text;
                    Thread th = new Thread(Thread_beginTest);
                    Thread th_log = new Thread(thread_log);
                    if (this.serialPort1.IsOpen && this.button5.Text == "开始测试")
                    {                        
                        this.splitContainer1.Panel1Collapsed = true; //隐藏平时打印窗口
                        this.splitContainer1.Panel2Collapsed = false;//显示测试打印区域
                        test_time = 1;  //当前循环次数
                        this.button5.Text = "停止测试";
                        this.button7.Enabled = false; //测试期间，发送按钮不可用
                        canstop_test = false;
                        canstop_log = false;
                        com_content = string.Empty;
                        List<string> lt = new List<string>(new string[] { "测试次数", "用例名称", "测试用例", "转换报文", "预计接收", "转换预计接收", "实际接收", "测试结果" });
                        this.listView1.Items.Clear();
                        this.listView1.Columns.Clear();
                        Listview_Columns_add(this.listView1, lt);
                        th.Start();                        
                        th_log.Start();
                    }
                    else if (this.button5.Text == "停止测试")
                    {
                        this.button5.Text = "开始测试";
                        this.button7.Enabled = true;
                        canstop_test = true;//线程th停止    
                        canstop_log = true;//线程th_log停止
                        th.Abort();//释放线程th资源
                        th_log.Abort();
                    }
                    else if (this.button5.Text == "完成测试")
                    {                                        
                        this.button5.Text = "开始测试";
                        this.button7.Enabled = true;
                        canstop_test = true;//线程th停止    
                        canstop_log = true;//线程th_log停止
                        th.Abort();//释放线程th资源
                        th_log.Abort();
                    }
                    else
                    {
                        canstop_test = true;//线程th停止    
                        canstop_log = true;//线程th_log停止
                        th.Abort();//释放线程th资源
                        th_log.Abort();
                        MessageBox.Show("串口未打开", "提示", MessageBoxButtons.OK);
                    }
                }
                catch
                {
                    MessageBox.Show("输入非数字", "提示", MessageBoxButtons.OK);
                }
            }
            else
                MessageBox.Show("请检查文件路径", "提示", MessageBoxButtons.OK);
        }

        /// <summary>
        /// 打开文件按钮：点击后选择用例文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {          
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "请选择文件";
            this.openFileDialog1.Filter = "Excel文件|*.xls";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox4.Text = this.openFileDialog1.FileName;
            }               
        }

        /// <summary>
        /// 生成用例按钮：点击后把用例转换成报文格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            filepath = this.textBox4.Text;            
            Thread th = new Thread(Thread_testcase); 
            if (this.button4.Text == "生成用例")
            {
                this.button4.Text = "点击结束";
                this.progressBar1.Value = 0; //每次开始都是重新开始
                canstop_case = false;
                th.Start();//生成用例的线程开始
            }
            else if (this.button4.Text == "点击结束")
            {
                this.button4.Text = "生成用例";
                canstop_case = true; //生成用例的线程停止标志位true：停止
                th.Abort();//生成用例的线程资源释放
            }
        }

        /// <summary>
        /// comboBox1键盘按键点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //当点击del时候删除当前值
            if (e.KeyChar == 46)
            {                
                comboBox1.Items.RemoveAt(comboBox1.SelectedIndex);
            }
        }

        /// <summary>
        /// textBox2键盘按键点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //当点击enter按键的时候
            if (e.KeyChar == 13)
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    if (this.textBox2.Text != null && this.textBox2.Text != "")
                    {
                        comboBox1.Items.Add(this.textBox2.Text);
                    }
                }
                else
                {
                    if (textBox2.Text != comboBox1.Text && textBox2.Text != null && textBox2.Text != "" && comboBox1.Items.Count > 1)
                        comboBox1.Items[comboBox1.SelectedIndex] = textBox2.Text;
                    if (textBox2.Text == "")
                    {
                        comboBox1.Items.RemoveAt(comboBox1.SelectedIndex);
                        comboBox1.SelectedIndex = 0;
                    }                       
                }
            }           
        }

        private int time, num_times;//初始化循环次数和发送间隔时间
        private List<string> send_content = new List<string>(); //循环发送的列表
        private bool canstop = false;//循环发送的线程停止标志位

        /// <summary>
        /// 新建的线程：循环发送方法
        /// </summary>
        private void send_new()
        {
            
            for (int i=0;i< num_times;i++)
            {
                //Console.WriteLine("i:" + i);
                foreach (string content in send_content)
                {
                    if (!canstop)
                    {
                        Write_data(content);
                        //Console.WriteLine("content:" + content);
                        Thread.Sleep(time);                        
                    }
                    else
                        break;
                }
                if (canstop)
                    break;
            }
            //循环完成后托管button7显示发送
            if (!canstop)
            {
                this.BeginInvoke(new MethodInvoker(delegate
                {
                    this.button7.Text = "发送";
                }));
            }                     
        }
        
        /// <summary>
        /// 鼠标双击菜单栏事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuStrip1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //如果平时打印区域显示，鼠标双击后如下
            if (this.splitContainer1.Panel1Collapsed == false)
            {
                this.splitContainer1.Panel1Collapsed = true;//平时打印区域隐藏
                this.splitContainer1.Panel2Collapsed = false;//测试打印区域显示
            }
            //如果平时打印区域隐藏，鼠标双击后如下
            else
            {
                this.splitContainer1.Panel1Collapsed = false;//平时打印区域显示
                this.splitContainer1.Panel2Collapsed = true;//测试打印区域隐藏
            }
        }

        /// <summary>
        /// 日常打印的窗口有更新时候，滚动条始终在最后一行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            richTextBox2.SelectionStart = richTextBox2.TextLength;
            richTextBox2.ScrollToCaret();
        }

        /// <summary>
        /// 右边的 发送按钮：点击后循环发送comboBox1内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                time = Convert.ToInt32(this.textBox6.Text);
                num_times = Convert.ToInt32(this.textBox5.Text);
                //Console.WriteLine("time:" + time);
                //Console.WriteLine("num_times:" + num_times);
                
                Thread th = new Thread(send_new);//新建线程th，send_new添加进线程执行中
                if (this.serialPort1.IsOpen && this.button7.Text == "发送")
                {
                    this.button7.Text = "停止";
                    this.button5.Enabled = false;//发送期间，开始测试按钮不可用
                    canstop = false;
                    send_content.Clear();//开始先清除发送列表

                    //除了第一个项外，挨个把其余项添加进发送列表send_content里
                    foreach (string text in comboBox1.Items)
                    {
                        if (text != "请添加...")
                        {
                            send_content.Add(text);
                        }                       
                    }
                    th.Start();  //线程th开始运行                  
                }
                else if (this.button7.Text == "停止")
                {
                    this.button7.Text = "发送";
                    this.button5.Enabled = true;
                    canstop = true;//线程th停止             
                    th.Abort();//释放线程th资源
                }
                else
                {
                    canstop = true;
                    th.Abort();
                    MessageBox.Show("串口未打开", "提示", MessageBoxButtons.OK);
                }
            }
            catch
            {
                MessageBox.Show("输入非数字", "提示", MessageBoxButtons.OK);
            }
        }
    }
}