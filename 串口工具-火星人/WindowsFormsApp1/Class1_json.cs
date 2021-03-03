using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO.Ports;
using System.IO;
using System.Runtime.CompilerServices;

namespace WindowsFormsApp1
{
    class Class1_json
    {
        /// <summary>
        /// 保存方式通用格式
        /// </summary>
        public class General_format
        {
            public string strOtherOne { get; set; }
            public string strOtherTwo { get; set; }
            public string strOtherThree { get; set; }
        }

        /// <summary>
        /// Listview数据保存格式
        /// </summary>
        public class List_format
        {
            public string one { get; set; }
            public string two { get; set; }
            public string three { get; set; }
            public string four { get; set; }
            public string five { get; set; }
            public string six { get; set; }
            public string seven { get; set; }
            public string eight { get; set; }
            public string nine { get; set; }
            public string ten { get; set; }

        }

        /// <summary>
        /// 串口配置
        /// </summary>
        public class Serial_config
        {
            /// <summary>
            /// 波特率
            /// </summary>
            public string intBaudRate { get; set; }
            /// <summary>
            /// 数据位：一般默认8
            /// </summary>
            public string intDataBits { get; set; }
            /// <summary>
            /// 停止位
            /// </summary>
            public string intStopBits { get; set; }
            /// <summary>
            /// 奇偶校验
            /// </summary>
            public string intParity { get; set; }
        }

        /// <summary>
        /// 报文配置
        /// </summary>
        public class Command_config
        {
            public Command_config()
            {
                this.strFrameHead = new General_format();
                this.strFrameTail = new General_format();
                this.strFrameData = new General_format();
                this.strFrameCommand = new General_format();
                this.strFrameOne = new General_format();
                this.strFramTwo = new General_format();
                this.strFramThree = new General_format();
            }
            /// <summary>
            /// 数据帧头
            /// </summary>
            public General_format strFrameHead { get; set; }

            /// <summary>
            /// 数据帧尾
            /// </summary>
            public General_format strFrameTail { get; set; }

            /// <summary>
            /// CRC校验模式
            /// </summary>
            public string strFrameCheckCode { get; set; }

            /// <summary>
            /// 是序号所占字节数
            /// </summary>
            public string intFrameSequential { get; set; }

            /// <summary>
            /// 数据内容
            /// </summary>
            public General_format strFrameData { get; set; }

            public General_format strFrameCommand { get; set; }
            public General_format strFrameOne { get; set; }
            public General_format strFramTwo { get; set; }
            public General_format strFramThree { get; set; }
        }

        /// <summary>
        /// 反馈实体
        /// </summary>
        public class My_config
        {
            public My_config()
            {
                this.Command_config = new Command_config();
                this.Serial_config = new Serial_config();
                this.Command_Name = new General_format();
                this.List_Name = new List_format();
                this.List_Value = new List_format();
                this.CheckCode_value = new List_format();
            }
            /// <summary>
            /// 报文配置
            /// </summary>
            public Command_config Command_config { get; set; }
            /// <summary>
            /// 串口配置
            /// </summary>
            public Serial_config Serial_config { get; set; }

            /// <summary>
            /// 报文组成成员配置
            /// </summary>
            public General_format Command_Name { get; set; }
            public List_format List_Name { get; set; }
            public List_format List_Value { get; set; }
            public List_format CheckCode_value { get; set; }
        }

        public static string srpath = @".\Config_json.json";

        /// <summary>
        /// 读取json文件
        /// </summary>
        /// <returns>成功则返回My_config类型的文件内容</returns>
        public static My_config ReadFile()
        {
            string reader;
            My_config my_config;
            try
            {
                //打开当前路径下的Config_json.json文件
                StreamReader sr = File.OpenText(srpath);
                reader = sr.ReadToEnd();
                sr.Close();
            }
            catch
            {
                reader = "";
            }
            my_config = JsonConvert.DeserializeObject<My_config>(reader);
            return my_config;
        }

        /// <summary>
        /// 写入数据到当前路径的Config_json.json文件里
        /// </summary>
        /// <param name="my_config">符合固定模型的数据</param>
        /// <returns>成功写入返回true</returns>
        public static bool WriteFile(My_config my_config)
        {
            string reader = JsonConvert.SerializeObject(my_config);
            //Console.WriteLine("hahaha:" + reader);
            try
            {
                File.WriteAllText(srpath, reader, Encoding.UTF8);
                return true;
            }
            catch(Exception e)
            {
                Console.WriteLine("失败原因:" + e.Message);
                return false;
            }
        }
    }
}