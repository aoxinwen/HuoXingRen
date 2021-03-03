using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static WindowsFormsApp1.Class1_json;

namespace WindowsFormsApp1
{
    public partial class Form2_messageSET : Form
    {
        public Form2_messageSET()
        {
            InitializeComponent();
        }
                                                                             
        //public static string[] com1Text = { "请输入..", "帧头", "时序", "命令码", "数据", "校验", "帧尾" }; //默认选项名
        public static string[] FrameSequential = { "请输入..", "0", "1", "2", "3" };                          //是序号的位数
        public static string[] FrameCheckCode = { "crc-8" , "crc-8-darc" , "crc-8-i-code" , "crc-8-itu" ,
        "crc-8-maxim" , "crc-8-rohc" , "crc-8-wcdma" ,"crc-16","crc-16-buypass","crc-16-dds-110",
        "crc-16-dect","crc-16-dnp","crc-16-en-13757","crc-16-genibus","crc-16-maxim","crc-16-mcrf4xx",
        "crc-16-riello","crc-16-t10-dif","crc-16-teledisk","crc-16-usb","x-25","xmodem","modbus",
        "kermit","crc-ccitt-false","crc-aug-ccitt","crc-24","crc-24-flexray-a","crc-24-flexray-b","crc-32",
        "crc-32-bzip2","crc-32c","crc-32d","crc-32-mpeg","posix","crc-32q","jamcrc","xfer","无校验"};         //crc校验模式
        //public static List<string> _list = new List<string>(com1Text);                                      //默认选项名的列表
        public static List<string> list_FrameHead = new List<string>();                                       //帧头 选项名的列表
        public static List<string> list_FrameTail = new List<string>();                                       //帧尾 选项名的列表
        public static List<string> list_FrameData = new List<string>();                                       //数据 选项名的列表
        public static List<string> list_FrameCommand = new List<string>();                                    //命令码 选项名的列表
        public static List<string> list_FrameOne = new List<string>();                                        //预留1 选项名的列表
        public static List<string> list_FrameTwo = new List<string>();                                        //预留2 选项名的列表
        public static List<string> list_FramThree = new List<string>();                                       //预留3 选项名的列表
        public static List<string> list_name = new List<string>();                                            //选项名称的列表        
        public static List<string> listview_name = new List<string>();                                        //listview里面显示的表头
        public static List<string> listview_value = new List<string>();                                       //listview里面显示的内容
        public static List<string> listview_check_value = new List<string>();                                 //listview里面内容校验的项



        /// <summary>
        /// 检查元素是否是列表的成员,是则返回true
        /// </summary>
        /// <param name="element">被检查的元素</param>
        /// <param name="aggregate">被检查的列表</param>
        /// <returns></returns>
        private static bool Is_include(string element, List<string> aggregate)
        {
            bool _result = false;
            foreach (string i in aggregate)
            {
                if (element == i)
                {
                    _result = true;
                    break;
                }
            }
            return _result;
        }

        /// <summary>
        /// 处理ComboBox.Text如果不在box.Items里面，则添加进入去（不足三个则依次添加，达到三个则替换第一个）
        /// </summary>
        /// <param name="box">ComboBox控件</param>
        /// <param name="general_">General_format类型的数据</param>
        /// <param name="list">暂存数据的list</param>
        /// <param name="start">第几位开始，默认0</param>
        /// <returns></returns>
        private static General_format Add_config_cmb(ComboBox box, General_format general_, List<string> list,int start = 0)
        {
            //如果输入的值不在常用值列表里面
            if (!Is_include(box.Text, list))
            //if (!box.Items.Contains(box.Text))
            {                
                switch (box.Items.Count - start)
                {                    
                    //还没有常用值的时候，添加第一个常用 数值
                    case 1:
                        list[start + 1] = box.Text;                        
                        break;
                    //只有1个常用值的时候，添加第2个常用 数值
                    case 2:
                        list[start + 2] = box.Text;                        
                        break;
                    //有2个常用值的时候，添加第3个常用 数值
                    case 3:
                        list[start + 3] = box.Text;                        
                        break;
                    //有3个常用值的时候，替换第1个常用 数值
                    case 4:
                        list[start + 1] = list[start + 2];
                        list[start + 2] = list[start + 3];
                        list[start + 3] = box.Text;                        
                        break;
                }
            }
            general_ = new General_format
            {
                strOtherOne = list[start + 1],
                strOtherTwo = list[start + 2],
                strOtherThree = list[start + 3]
            };
            Combox_add(box, list);
            box.SelectedIndex = 0;
            return general_;
        }

        /// <summary>
        /// 把list非null元素挨个添加进box的items
        /// </summary>
        /// <param name="box">要添加的box</param>
        /// <param name="list">被添加的list</param>
        private static void Combox_add(ComboBox box, List<string> list)
        {
            box.Items.Clear();
            foreach(string li in list)
            {
                if (li != null)
                {
                    box.Items.Add(li);
                }
            }
        }

        /// <summary>
        /// 把list的非null元素挨个添加进listview的Columns
        /// </summary>
        /// <param name="lv">要添加的ListView</param>
        /// <param name="list">被添加的list</param>
        private static void Listview_Columns_add(ListView lv,List<string> list)
        {
            lv.Clear();
            lv.BeginUpdate();
            foreach (string li in list)
            {
                if(li != null)
                {
                    lv.Columns.Add(li,90);
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
            lv.Items.Clear();
            ListViewItem lvi = new ListViewItem();
            lvi.UseItemStyleForSubItems = false;  //设置单元格可以修改颜色和格式
            lv.BeginUpdate();
            lvi.ImageIndex = 0;
            lvi.Text = list[0];
            for (int i = 1;i<list.Count;i++)
            {
                if (list[i] != null)
                {
                    lvi.SubItems.Add(list[i]);
                }
            }
            lv.Items.Add(lvi);
            lv.EndUpdate();
        }

        /// <summary>
        /// 返回添加表值后的List_format，并且listview显示表头
        /// </summary>
        /// <param name="lv"></param>
        /// <param name="list_Format"></param>
        /// <param name="list"></param>
        /// <param name="box"></param>
        /// <returns></returns>
        private static List_format Add_config_Columns_listView(ListView lv, List_format list_Format, List<string> list, ComboBox box)
        {
            switch (lv.Columns.Count)
            {
                case 0:
                    list[0] = box.Text;
                    break;
                case 1:
                    list[1] = box.Text;
                    break;
                case 2:
                    list[2] = box.Text;
                    break;
                case 3:
                    list[3] = box.Text;
                    break;
                case 4:
                    list[4] = box.Text;
                    break;
                case 5:
                    list[5] = box.Text;
                    break;
                case 6:
                    list[6] = box.Text;
                    break;
                case 7:
                    list[7] = box.Text;
                    break;
                case 8:
                    list[8] = box.Text;
                    break;
                case 9:
                    list[9] = box.Text;
                    break;
                case 10:
                    list[10] = box.Text;
                    break;
            }
            Listview_Columns_add(lv, list);
            list_Format = new List_format
            {
                one = list[1],
                two = list[2],
                three = list[3],
                four = list[4],
                five = list[5],
                six = list[6],
                seven = list[7],
                eight = list[8],
                nine = list[9],
                ten = list[10]
            };
            return list_Format;
        }

        /// <summary>
        /// 返回添加表值后的List_format，并且listview显示表值
        /// </summary>
        /// <param name="lv">要处理的listview</param>
        /// <param name="list_Format">被返回的数据结构</param>
        /// <param name="list">被添加的list</param>
        /// <param name="box">被处理的combox</param>
        /// <returns></returns>
        private static List_format Add_config_Value_listView(ListView lv, List_format list_Format, List<string> list, ComboBox box)
        {
            switch (lv.Columns.Count)
            {
                case 1:
                    list[0] = box.Text;
                    break;
                case 2:
                    list[1] = box.Text;
                    break;
                case 3:
                    list[2] = box.Text;
                    break;
                case 4:
                    list[3] = box.Text;
                    break;
                case 5:
                    list[4] = box.Text;
                    break;
                case 6:
                    list[5] = box.Text;
                    break;
                case 7:
                    list[6] = box.Text;
                    break;
                case 8:
                    list[7] = box.Text;
                    break;
                case 9:
                    list[8] = box.Text;
                    break;
                case 10:
                    list[9] = box.Text;
                    break;
                case 11:
                    list[10] = box.Text;
                    break;
            }
            Listview_value_add(lv, list);
            list_Format = new List_format
            {
                one = list[1],
                two = list[2],
                three = list[3],
                four = list[4],
                five = list[5],
                six = list[6],
                seven = list[7],
                eight = list[8],
                nine = list[9],
                ten = list[10]
            };
            return list_Format;
        }

        /// <summary>
        /// 返回选择校验项的List_format，并且listview显示选择的颜色
        /// </summary>
        /// <param name="lv"></param>
        /// <param name="list_Format"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        private static List_format Add_check_value(ListView lv, List_format list_Format, List<string> list, int i)
        {
            list[i - 1] = lv.Items[0].SubItems[i].Text;
            list_Format = new List_format
            {
                one = list[0],
                two = list[1],
                three = list[2],
                four = list[3],
                five = list[4],
                six = list[5],
                seven = list[6],
                eight = list[7],
                nine = list[8],
                ten = list[9]
            };
            return list_Format;
        }

        /// <summary>
        /// 返回删除后的list_Format，并且listview显示删除后的表头
        /// </summary>
        /// <param name="lv">要处理的listview</param>
        /// <param name="list_Format">被返回的数据结构</param>
        /// <param name="list">被删除的list</param>
        /// <returns></returns>
        private static List_format Clean_config_Columns_listView(ListView lv, List_format list_Format, List<string> list)
        {
            switch (lv.Columns.Count)
            {
                case 11:
                    list[10] = null;
                    break;
                case 10:
                    list[9] = null;
                    break;
                case 9:
                    list[8] = null;
                    break;
                case 8:
                    list[7] = null;
                    break;
                case 7:
                    list[6] = null;
                    break;
                case 6:
                    list[5] = null;
                    break;
                case 5:
                    list[4] = null;
                    break;
                case 4:
                    list[3] = null;
                    break;
                case 3:
                    list[2] = null;
                    break;
                case 2:
                    list[1] = null;
                    break;
                case 1:
                    break;
            }
            Listview_Columns_add(lv, list);
            list_Format = new List_format
            {
                one = list[1],
                two = list[2],
                three = list[3],
                four = list[4],
                five = list[5],
                six = list[6],
                seven = list[7],
                eight = list[8],
                nine = list[9],
                ten = list[10]
            };
            return list_Format;
        }

        /// <summary>
        /// 返回删除后的list_Format，并且listview显示删除后的表值
        /// </summary>
        /// <param name="lv">要处理的listview</param>
        /// <param name="list_Format">被返回的数据结构</param>
        /// <param name="list">被删除的list</param>
        /// <returns></returns>
        private static List_format Clean_config_Value_listView(ListView lv, List_format list_Format, List<string> list)
        {
            switch (lv.Columns.Count)
            {
                case 10:
                    list[10] = null;
                    break;
                case 9:
                    list[9] = null;
                    break;
                case 8:
                    list[8] = null;
                    break;
                case 7:
                    list[7] = null;
                    break;
                case 6:
                    list[6] = null;
                    break;
                case 5:
                    list[5] = null;
                    break;
                case 4:
                    list[4] = null;
                    break;
                case 3:
                    list[3] = null;
                    break;
                case 2:
                    list[2] = null;
                    break;
                case 1:
                    list[1] = null;
                    break;
                case 0:
                    break;
            }
            Listview_value_add(lv, list);
            list_Format = new List_format
            {
                one = list[1],
                two = list[2],
                three = list[3],
                four = list[4],
                five = list[5],
                six = list[6],
                seven = list[7],
                eight = list[8],
                nine = list[9],
                ten = list[10]
            };
            return list_Format;
        }

        /// <summary>
        /// 返回删除校验项的List_format，并且listview显示选择的颜色
        /// </summary>
        /// <param name="lv"></param>
        /// <param name="list_Format"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        private static List_format Clean_check_value(ListView lv, List_format list_Format, List<string> list)
        {
            switch (lv.Columns.Count)
            {
                case 10:
                    list[9] = null;
                    break;
                case 9:
                    list[8] = null;
                    break;
                case 8:
                    list[7] = null;
                    break;
                case 7:
                    list[6] = null;
                    break;
                case 6:
                    list[5] = null;
                    break;
                case 5:
                    list[4] = null;
                    break;
                case 4:
                    list[3] = null;
                    break;
                case 3:
                    list[2] = null;
                    break;
                case 2:
                    list[1] = null;
                    break;
                case 1:
                    list[0] = null;
                    break;
                case 0:
                    break;
            }
            list_Format = new List_format
            {
                one = list[0],
                two = list[1],
                three = list[2],
                four = list[3],
                five = list[4],
                six = list[5],
                seven = list[6],
                eight = list[7],
                nine = list[8],
                ten = list[9]
            };
            return list_Format;
        }

        /// <summary>
        /// 第i项从json删除
        /// </summary>
        /// <param name="lv"></param>
        /// <param name="list_Format"></param>
        /// <param name="list"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private static List_format white_check_value(ListView lv, List_format list_Format, List<string> list, int i)
        {
            list[i - 1] = null;
            list_Format = new List_format
            {
                one = list[0],
                two = list[1],
                three = list[2],
                four = list[3],
                five = list[4],
                six = list[5],
                seven = list[6],
                eight = list[7],
                nine = list[8],
                ten = list[9]
            };
            return list_Format;
        }

        /// <summary>
        /// 从json配置文件中导入上次默认配置
        /// </summary>
        /// <param name="my_Config">获取json配置的对象</param>
        private void Json_load()
        {
            My_config my_Config = ReadFile();
            //load 选项名
            list_name = new List<string>(new string[] { "请输入..", "帧头", "时序", "命令码", "数据", "校验", "帧尾",my_Config.Command_Name.strOtherOne, my_Config.Command_Name.strOtherTwo, my_Config.Command_Name.strOtherThree });
            /*_list = _list.Concat(list_name).ToList(); */
            Combox_add(this.comboBox1, list_name);
            this.comboBox1.SelectedIndex = 0;

            //load 各个选项的值
            //帧头：
            //list_FrameHead.Add("请输入..");
            list_FrameHead = new List<string>(new string[] { "请输入.." ,my_Config.Command_config.strFrameHead.strOtherOne, my_Config.Command_config.strFrameHead.strOtherTwo, my_Config.Command_config.strFrameHead.strOtherThree });           
            
            //帧尾
            //list_FrameTail.Add("请输入..");
            list_FrameTail = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFrameTail.strOtherOne, my_Config.Command_config.strFrameTail.strOtherTwo, my_Config.Command_config.strFrameTail.strOtherThree });

            //数据
            //list_FrameData.Add("请输入..");
            list_FrameData = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFrameData.strOtherOne, my_Config.Command_config.strFrameData.strOtherTwo, my_Config.Command_config.strFrameData.strOtherThree });

            //命令码
            //list_FrameCommand.Add("请输入..");
            list_FrameCommand = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFrameCommand.strOtherOne, my_Config.Command_config.strFrameCommand.strOtherTwo, my_Config.Command_config.strFrameCommand.strOtherThree });

            //预留1
            //list_FrameOne.Add("请输入..");
            list_FrameOne = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFrameOne.strOtherOne, my_Config.Command_config.strFrameOne.strOtherTwo, my_Config.Command_config.strFrameOne.strOtherThree });
            
            //预留2
            //list_FrameTwo.Add("请输入..");
            list_FrameTwo = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFramTwo.strOtherOne, my_Config.Command_config.strFramTwo.strOtherTwo, my_Config.Command_config.strFramTwo.strOtherThree });

            //预留3
            //list_FramThree.Add("请输入..");
            list_FramThree = new List<string>(new string[] { "请输入..", my_Config.Command_config.strFramThree.strOtherOne, my_Config.Command_config.strFramThree.strOtherTwo, my_Config.Command_config.strFramThree.strOtherThree });

            //listview表头名
            listview_name = new List<string>(new string[] { "报文名：",my_Config.List_Name.one, my_Config.List_Name.two, my_Config.List_Name.three, my_Config.List_Name.four, my_Config.List_Name.five, my_Config.List_Name.six, my_Config.List_Name.seven, my_Config.List_Name.eight, my_Config.List_Name.nine, my_Config.List_Name.ten });

            //listview表内容
            listview_value = new List<string>(new string[] { "数据值：",my_Config.List_Value.one, my_Config.List_Value.two, my_Config.List_Value.three, my_Config.List_Value.four, my_Config.List_Value.five, my_Config.List_Value.six, my_Config.List_Value.seven, my_Config.List_Value.eight, my_Config.List_Value.nine, my_Config.List_Value.ten });
            if (listview_name != null)
                Listview_Columns_add(this.listView1, listview_name);
            if (listview_value != null)
                Listview_value_add(this.listView1, listview_value);
            listview_check_value = new List<string>(new string[] { my_Config.CheckCode_value.one, my_Config.CheckCode_value.two, my_Config.CheckCode_value.three, my_Config.CheckCode_value.four, my_Config.CheckCode_value.five, my_Config.CheckCode_value.six, my_Config.CheckCode_value.seven, my_Config.CheckCode_value.eight, my_Config.CheckCode_value.nine, my_Config.CheckCode_value.ten });
            for (int i=0; i< listview_check_value.Count;i++)
            {
                if (listview_check_value[i] != null)
                    this.listView1.Items[0].SubItems[i + 1].BackColor = Color.Aqua;
            }
        }

        /// <summary>
        /// 根据选项名字来显示当前选项的默认值
        /// </summary>
        private void Default_value()
        {
            switch (this.comboBox1.SelectedIndex)
            {
                case 0:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    this.comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
                    break;
                case 1:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameHead);
                    /*this.comboBox2.Items.AddRange(list_FrameHead.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 2:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    this.comboBox2.Items.AddRange(FrameSequential);
                    this.comboBox2.SelectedIndex = 2;
                    break;

                case 3:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameCommand);
                    /*this.comboBox2.Items.AddRange(list_FrameCommand.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 4:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameData);
                    /*this.comboBox2.Items.AddRange(list_FrameData.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 5:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    this.comboBox2.Items.AddRange(FrameCheckCode);
                    this.comboBox2.SelectedIndex = 14;
                    break;
                case 6:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameTail);
                    /*this.comboBox2.Items.AddRange(list_FrameTail.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 7:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameOne);
                    /*this.comboBox2.Items.AddRange(list_FrameOne.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 8:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FrameTwo);
                    /*this.comboBox2.Items.AddRange(list_FrameTwo.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
                case 9:
                    this.comboBox2.Items.Clear();
                    this.comboBox2.Text = null;
                    Combox_add(this.comboBox2, list_FramThree);
                    /*this.comboBox2.Items.AddRange(list_FramThree.ToArray());*/
                    this.comboBox2.SelectedIndex = 0;
                    break;
            }
        }

        /// <summary>
        /// 把新配置写入json文件中
        /// </summary>
        private void Json_write()
        {
            My_config my_Config = ReadFile();
            var config = new My_config();
            var command_config = new Command_config();
            var FrameHead = new General_format();
            var FrameTail = new General_format();
            var FrameData = new General_format();
            var FrameCommand = new General_format();
            var FrameOne = new General_format();
            var FramTwo = new General_format();
            var FramThree = new General_format();
            var command_name = new General_format();
            var list_Name = new List_format();
            var list_Value = new List_format();

            //listview 列表表头
            list_Name = Add_config_Columns_listView(this.listView1, list_Name, listview_name,this.comboBox1);

            //listview 列表表值
            list_Value = Add_config_Value_listView(this.listView1, list_Value, listview_value,this.comboBox2);

            //选项名
            //根据实际个数来写入配置:选项名称 中
            switch (comboBox1.SelectedIndex)
            {
                //选择“帧头”选项时
                case 1:
                    FrameHead = Add_config_cmb(comboBox2, FrameHead, list_FrameHead);
                    command_config = new Command_config
                    {
                        strFrameHead = FrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                    };
                    break;
                //选择“时序”的时候，显示框内容就是时序位
                case 2:
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = this.comboBox2.Text,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                    };
                    break;
                //选择“命令码”的时候
                case 3:
                    FrameCommand = Add_config_cmb(comboBox2, FrameCommand, list_FrameCommand);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = FrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“数据”的时候
                case 4:
                    FrameData = Add_config_cmb(comboBox2, FrameData, list_FrameData);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = FrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“校验”的时候
                case 5:
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = comboBox2.Text,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“帧尾”的时候
                case 6:
                    FrameTail = Add_config_cmb(comboBox2, FrameTail, list_FrameTail);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = FrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“预留1”的时候
                case 7:
                    FrameOne = Add_config_cmb(comboBox2, FrameOne, list_FrameOne);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = FrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“预留2”的时候
                case 8:
                    FramTwo = Add_config_cmb(comboBox2, FramTwo, list_FrameTwo);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = FramTwo,
                        strFramThree = my_Config.Command_config.strFramThree
                        
                    };
                    break;
                //选择“预留3”的时候
                case 9:
                    FramThree = Add_config_cmb(comboBox2, FramThree, list_FramThree);
                    command_config = new Command_config
                    {
                        strFrameHead = my_Config.Command_config.strFrameHead,
                        strFrameTail = my_Config.Command_config.strFrameTail,
                        strFrameCheckCode = my_Config.Command_config.strFrameCheckCode,
                        intFrameSequential = my_Config.Command_config.intFrameSequential,
                        strFrameData = my_Config.Command_config.strFrameData,
                        strFrameCommand = my_Config.Command_config.strFrameCommand,
                        strFrameOne = my_Config.Command_config.strFrameOne,
                        strFramTwo = my_Config.Command_config.strFramTwo,
                        strFramThree = FramThree

                    };
                    break;
            }
            //选项名称
            command_name = Add_config_cmb(comboBox1, command_name, list_name,6);

            //赋值config
            config = new My_config
            {
                Command_config = command_config,
                Serial_config = my_Config.Serial_config,
                Command_Name = command_name,
                List_Name = list_Name,
                List_Value = list_Value,
                CheckCode_value = my_Config.CheckCode_value
            };
                        
            //把config烧写进json文件
            if(!WriteFile(config)) //调用WriteFile方法把配置写入json文件中
            {
                MessageBox.Show("保存失败", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }            

            //写完数据后显示已经标记的颜色
            for (int i = 0; i < listview_check_value.Count; i++)
            {
                if (listview_check_value[i] != null)
                    this.listView1.Items[0].SubItems[i + 1].BackColor = Color.Aqua;
            }
        }

        /// <summary>
        /// 颜色选择水绿色后把校验项写入json
        /// </summary>
        /// <param name="i"></param>
        private void Json_write_check_value(int i)
        {
            My_config my_Config = ReadFile();
            var config = new My_config();
            var checkCode_value = new List_format();

            checkCode_value = Add_check_value(this.listView1, checkCode_value, listview_check_value, i);

            //赋值config
            config = new My_config
            {
                Command_config = my_Config.Command_config,
                Serial_config = my_Config.Serial_config,
                Command_Name = my_Config.Command_Name,
                List_Name = my_Config.List_Name,
                List_Value = my_Config.List_Value,
                CheckCode_value = checkCode_value
            };
            //把config烧写进json文件
            WriteFile(config); //调用WriteFile方法把配置写入json文件中
        }

        /// <summary>
        /// 删除最新的1列数据并保存在json中
        /// </summary>
        private void Json_clean()

        {
            My_config my_Config = ReadFile();
            var config = new My_config();
            var list_Name = new List_format();
            var list_Value = new List_format();
            var checkCode_value = new List_format();

            //listview 列表表头
            list_Name = Clean_config_Columns_listView(this.listView1, list_Name, listview_name);
            //listview 列表表值
            list_Value = Clean_config_Value_listView(this.listView1, list_Value, listview_value);

            //校验内容
            checkCode_value = Clean_check_value(this.listView1, checkCode_value, listview_check_value);

            //赋值config
            config = new My_config
            {
                Command_config = my_Config.Command_config,
                Serial_config = my_Config.Serial_config,
                Command_Name = my_Config.Command_Name,
                List_Name = list_Name,
                List_Value = list_Value,
                CheckCode_value = checkCode_value
            };
            //把config烧写进json文件
            WriteFile(config); //调用WriteFile方法把配置写入json文件中
            for (int i = 0; i < listview_check_value.Count; i++)
            {
                if (listview_check_value[i] != null)
                    this.listView1.Items[0].SubItems[i + 1].BackColor = Color.Aqua;
            }
        }

        /// <summary>
        /// 颜色选择白色后把校验项从json删除
        /// </summary>
        /// <param name="i"></param>
        private void Json_clean_check_value(int i)
        {
            My_config my_Config = ReadFile();
            var config = new My_config();
            var checkCode_value = new List_format();
            //校验内容
            checkCode_value = white_check_value(this.listView1, checkCode_value, listview_check_value, i);

            //赋值config
            config = new My_config
            {
                Command_config = my_Config.Command_config,
                Serial_config = my_Config.Serial_config,
                Command_Name = my_Config.Command_Name,
                List_Name = my_Config.List_Name,
                List_Value = my_Config.List_Value,
                CheckCode_value = checkCode_value
            };
            //把config烧写进json文件
            WriteFile(config); //调用WriteFile方法把配置写入json文件中
        }

        /// <summary>
        /// 自动加载的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_messageSET_Load(object sender, EventArgs e)
        {
            this.Text = "报文设置";
            this.listView1.View = View.Details;
            this.listView1.SmallImageList = this.imageList1;
            Json_load();

            //this.listView1.Items[0].SubItems[0].BackColor = 

        }

        /// <summary>
        /// 选项名改变为：“请输入..”后该comboBox可以手动输入，否则不能输入只能选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //选择“请输入..”后该comboBox可以手动输入，否则不能输入只能选择
            if (this.comboBox1.Text == "请输入..")
            {
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
            }
            else
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            //根据选项名字来显示当前选项的默认值
            Default_value();
        }

        /// <summary>
        /// 选项值改变为：“请输入..”后该comboBox可以手动输入，否则不能输入只能选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //选择“请输入..”后该comboBox可以手动输入，否则不能输入只能选择
            if (this.comboBox2.Text == "请输入..")
            {
                this.comboBox2.DropDownStyle = ComboBoxStyle.DropDown;
            }
            else
                this.comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        /// <summary>
        /// 提交按钮：点击后确认最终报文的配置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Json_write();
        }

        /// <summary>
        /// 清除按钮：点击一次删除最后一个配置直到删除全部
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            Json_clean();
        }

        /// <summary>
        /// 确认按键：点击把报文格式推送到串口发送
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
              //点击确认按键把配置写入json文件
            this.DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// 取消按钮：点击后取消操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        /// <summary>
        /// 选择当前列后，值变成水绿色，再点击变成白色
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (this.listView1.Items[0].SubItems[e.Column].BackColor != Color.Aqua)
            {
                this.listView1.Items[0].SubItems[e.Column].BackColor = Color.Aqua;
                Json_write_check_value(e.Column);
            }                
            else
            {
                this.listView1.Items[0].SubItems[e.Column].BackColor = Color.White;
                Json_clean_check_value(e.Column);
            }
        }
    }
}
