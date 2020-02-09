/*----------------------------------------------------------------
// Copyright (C) 2007宁波金唐软件有限公司
//
// 文件名：全局方法
// 文件功能描述：
//
// 创建标识：
//
// 修改标识：朱增洪
// 修改描述：新增提示框方法
//           
// 修改标识：
// 修改描述：
//----------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.Security.Cryptography;

namespace ExcelToPdf
{
    /// <summary>
    /// 系统功能类
    /// </summary>
    public static class Sysfun
    {
        #region 数据类型转换

        /// <summary>
        /// 转值到DateTime
        /// </summary>
        public static DateTime ConverToDateTime(object obj, DateTime defaultvalue)
        {
            if (obj is System.DBNull || obj == null)
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToDateTime(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转值到Int
        /// </summary>
        public static short ConverToShort(object obj, short defaultvalue)
        {
            if (obj is System.DBNull || obj == null || !IsNum(obj.ToString()))
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToInt16(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转值到Int
        /// </summary>
        public static int ConverToInt(object obj, int defaultvalue)
        {
            if (obj is System.DBNull || obj == null || !IsNum(obj.ToString()))
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToInt32(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转值到Long
        /// </summary>
        public static long ConverToLong(object obj, long defaultvalue)
        {
            if (obj is System.DBNull || obj == null || !IsNum(obj.ToString()))
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToInt64(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转值到Decimal
        /// </summary>
        public static decimal ConverToDecimal(object obj, decimal defaultvalue)
        {
            if (obj is System.DBNull || obj == null || !IsNum(obj.ToString()) )
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToDecimal(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转到Double
        /// </summary>
        public static double ConverToDouble(object obj, double defaultvalue)
        {
            if (obj is System.DBNull || obj == null || !IsNum(obj.ToString()))
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToDouble(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 转到String
        /// </summary>
        public static string ConverToString(object obj, string defaultvalue)
        {
            if (obj is System.DBNull || obj == null)
                return defaultvalue;
            else
            {
                try
                {
                    return Convert.ToString(obj);
                }
                catch
                {
                    return defaultvalue;
                }
            }
        }

        /// <summary>
        /// 从字符串类型转换到颜色RGB格式
        /// </summary>
        /// <param name="color3str">颜色字符串(255,255,255)或255,255,255</param>
        /// <param name="defaultColor">默认颜色</param>
        /// <returns>RGB颜色</returns>
        public static Color ConverToColorRGB(string color3str, Color defaultColor)
        {
            color3str = color3str.Replace("(", "");
            color3str = color3str.Replace(")", "");
            string[] FColor = color3str.Split(',');

            if (FColor.Length != 3) 
                return defaultColor;
            else
            {
                Color color;
                try
                {
                    color = System.Drawing.Color.FromArgb(Convert.ToInt32(FColor[0]), Convert.ToInt32(FColor[1]), Convert.ToInt32(FColor[2]));
                }
                catch (System.Exception e)
                {
                    MessageBox.Show("颜色转换有误! " + e.Message);
                    return defaultColor;
                }
                return color;
            }
        }

        #endregion 

        #region 功能函数

        /// <summary>
        /// 判断字符串是否是数字(使用正则表达式检查)
        /// </summary>
        /// <param name="str"></param>
        /// <returns>true/false</returns>
        public static bool IsNumeric(string str)
        {
            if (str == null || str.Length == 0 || str == "-")
                return false;
            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"^[-]?\d*[.]?\d*$");  // 包括: 0.74, .74, 10, 10. , 10.1
            return reg.IsMatch(str);
        }

        /// <summary>
        /// 判断字符串是否是数字(也,使用正则表达式检查, 原来:只检验,0..9,+-.) 为兼容以前版本保留此函数
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNum(string str)
        {
            //if (str==null || str.Length == 0)
            //    return false;
            //int numCnt = 0;
            //foreach (char c in str)
            //    if (c >= '0' && c <= '9')
            //    {
            //        numCnt++;
            //        continue;
            //    }
            //    else if ("-+.".IndexOf(c) < 0)
            //        return false;
            //if (numCnt == 0)
            //    return false;
            //else
            //    return true;
            if (str == null || str.Length == 0 || str == "-")
                return false;
            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"^[-]?\d*[.]?\d*$");  // 包括: 0.74, .74, 10, 10. , 10.1
            return reg.IsMatch(str);
        }

        /// <summary>
        /// 判断字符串是否是整数
        /// </summary>
        /// <param name="str"></param>
        /// <returns>true/false</returns>
        public static bool IsInt(string str)
        {
            if (str == null || str.Length == 0)
                return false;
            foreach (char c in str)
                if (c >= '0' && c <= '9')
                    continue;
                else
                    return false;
            return true;
        }

        /// <summary>
        /// 获取月份的最后一天
        /// </summary>
        /// <param name="DayOfMonth">月份中的一天</param>
        /// <returns>该月的最后一天</returns>
        public static DateTime GetLastDateForMonth(DateTime DayOfMonth)
        {
            DateTime DtEnd ; 
            int Dtyear, DtMonth; 
            Dtyear = DayOfMonth.Year; 
            DtMonth = DayOfMonth.Month; 
            int MonthCount = DateTime.DaysInMonth(Dtyear, DtMonth);//计算该月有多少天 
            DtEnd = Convert.ToDateTime(Dtyear.ToString() + "-" + DtMonth.ToString() + "-" + MonthCount); 
            DtEnd.AddDays(1).AddMilliseconds(-1);
            return DtEnd;
        }


        /// <summary>
        /// 半角转全角
        /// </summary>
        /// <param name="strChar">转换的字符串</param>
        /// <returns></returns>
        public static string ConvertDbcToSbc(string strChar)
        {
            //char[] c = strChar.ToCharArray();
            //for (int i = 0; i < c.Length; i++)
            //{
            //    byte[] b = System.Text.Encoding.Unicode.GetBytes(c, i, 1);
            //    if (b.Length == 2)
            //    {
            //        if (b[1] == 0)
            //        {
            //            b[0] = (byte)(b[0] - 32);
            //            b[1] = 255;
            //            c[i] = System.Text.Encoding.Unicode.GetChars(b)[0];
            //        }
            //    }
            //}
            ////返回全角串
            //return new string(c);
            char[] c = strChar.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 32)
                {
                    c[i] = (char)12288;
                    continue;
                }
                if (c[i] < 127)
                    c[i] = (char)(c[i] + 65248);
            }
            return new string(c);
        }

        /// <summary>
        /// 全角转半角
        /// </summary>
        /// <param name="strChar"></param>
        /// <returns></returns>
        public static string ConvertSbcToDbc(string strChar)
        {
            //char[] c = strChar.ToCharArray();
            //for (int i = 0; i < c.Length; i++)
            //{
            //    byte[] b = System.Text.Encoding.Unicode.GetBytes(c, i, 1);
            //    if (b.Length == 2)
            //    {
            //        if (b[1] == 255)
            //        {
            //            b[0] = (byte)(b[0] + 32);
            //            b[1] = 0;
            //            c[i] = System.Text.Encoding.Unicode.GetChars(b)[0];
            //        }
            //    }
            //}

            ////返回半角串
            //return new string(c);
            strChar = strChar.Replace("％", "%");
            strChar = strChar.Replace("。", ".");
            strChar = strChar.Replace("，", ",");
            strChar = strChar.Replace("》", ">");
            strChar = strChar.Replace("《", "<");
            strChar = strChar.Replace("（", "(");
            strChar = strChar.Replace("）", ")");
            char[] c = strChar.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 12288)
                {
                    c[i] = (char)32;
                    continue;
                }
                if (c[i] > 65280 && c[i] < 65375)
                    c[i] = (char)(c[i] - 65248);
            }
            return new string(c);

        }

        /// <summary>
        /// 获取字符串在数据库的长度
        /// </summary>
        /// <param name="aOrgStr"></param>
        /// <returns></returns>
        public static int GetStringDBLength(String aOrgStr)
        {
            int intLen = aOrgStr.Length;
            int i;
            char[] chars = aOrgStr.ToCharArray();
            for (i = 0; i < chars.Length; i++)
            {
                if (System.Convert.ToInt32(chars[i]) > 255)
                {
                    intLen ++;
                }
            }
            return intLen;
        } 

        /// <summary>
        /// 增加SQL的Where条件
        /// </summary>
        /// <param name="originalSql">原始SQL</param>
        /// <param name="condition">加入的条件</param>
        /// <returns>新的SQL</returns>
        public static string SQLAddWhere(string originalSql, string condition)
        {
            //设置新的sqlselect
            string sql = originalSql.ToUpper();
            int li_locBef, li_locLst;

            // 从最后一个FROM后面查找WHERE条件
            li_locBef = sql.LastIndexOf("FROM ");
            li_locBef = sql.IndexOf("WHERE ", li_locBef);

            if (li_locBef < 0)  // 无WHERE
            {
                li_locBef = sql.IndexOf("GROUP BY ");
                if (li_locBef < 0)  // 无GROUP
                {
                    li_locBef = sql.IndexOf("ORDER BY ");
                    if (li_locBef < 0)  // 无ORDER
                    {
                        originalSql += " WHERE " + condition + " ";
                    }
                    else  // 有ORDER
                    {
                        originalSql = originalSql.Substring(0, li_locBef - 1) + " WHERE " + condition + originalSql.Substring(li_locBef - 1, originalSql.Length - li_locBef + 1);
                    }
                }
                else  // 有GROUP
                {
                    originalSql = originalSql.Substring(0, li_locBef - 1) + " WHERE " + condition + originalSql.Substring(li_locBef - 1, originalSql.Length - li_locBef + 1);
                }
            }
            else  // 有WHERE
            {
                li_locLst = sql.IndexOf("GROUP BY");
                if (li_locLst >= 0)  // 有GROUP
                {
                    originalSql = originalSql.Substring(0, li_locLst - 1) + " AND (" + condition + ") " + originalSql.Substring(li_locLst - 1, originalSql.Length - li_locLst + 1);
                }
                else
                {
                    li_locLst = sql.IndexOf("ORDER BY");
                    if (li_locLst >= 0)  // 有ORDER
                    {
                        originalSql = originalSql.Substring(0, li_locLst - 1) + " AND (" + condition + ") " + originalSql.Substring(li_locLst - 1, originalSql.Length - li_locLst + 1);
                    }
                    else
                    {
                        originalSql += " AND (" + condition + ") ";
                    }
                }
            }
            return originalSql;
        }

        /// <summary>
        /// 返回参数值列表 
        /// </summary>
        /// <param name="Para">多值字符串 (用;分隔)</param>
        /// <returns></returns>
        public static ArrayList GetListFromPara(string Para, string SplitStr)
        {
            ArrayList valuelist = new ArrayList();
            while (Para.Length > 0)
            {
                if (Para.IndexOf(SplitStr) >= 0)
                {
                    string Value = Para.Substring(0, Para.IndexOf(SplitStr));
                    if (!valuelist.Contains(Value))
                    {
                        valuelist.Add(Value);
                    }

                    Para = Para.Substring(Para.IndexOf(SplitStr) + 1);
                }
                else
                {
                    valuelist.Add(Para);
                    break;
                }

            }

            return valuelist;
        }

        /// <summary>
        /// 对象深度拷贝
        /// </summary>
        /// <param name="SourceObj">源对象</param>
        /// <param name="TargetObj">目标对象</param>
        public static void CopyTo(object SourceObj, object TargetObj)
        {
            System.Reflection.PropertyInfo[] propertiesSou = (SourceObj.GetType()).GetProperties();
            System.Reflection.PropertyInfo[] propertiesTar = (TargetObj.GetType()).GetProperties();
            for (int i = 0; i < propertiesSou.Length; i++)
            {
                propertiesTar[i].SetValue(TargetObj, propertiesSou[i].GetValue(SourceObj, null), null);
            }
        }

        /// <summary>
        ///  转换成大写金额
        /// </summary>
        /// <param name="money">金额</param>
        /// <returns>大写金额</returns>
        public static string UpperMoney(decimal money)
        {
            string upper_number_str = "零壹贰叁肆伍陆柒捌玖";
            string upper_unit_str = "十亿仟佰拾万仟佰拾元角分";
            string upper_number;           //  数值的大写
            string upper_unit;             //  单位大写
            string money_number;           //  金额数字
            string upper_money;            //  大写金额

            int number_len;
            int point_unit;
            string tmp_ch;
            int tmp_num;

            money = System.Decimal.Round(money, 2) * 100.00M;
            if (money == 0.00M)         //    零返回空字符串
                return "";

            money_number = money.ToString();
            point_unit = money_number.IndexOf('.');
            if (point_unit > 0)
                money_number = money_number.Substring(0, point_unit);
            number_len = money_number.Length;
            upper_money = "整";

            point_unit = upper_unit_str.Length - 1;
            upper_unit = "";
            upper_number = "";

            for (int i = number_len - 1; i >= 0; i--, point_unit--)
            {
                tmp_ch = money_number.Substring(i, 1);

                if ("9876543210 ".IndexOf(tmp_ch) >= 0)
                {
                    tmp_num = int.Parse(tmp_ch);
                    upper_number = upper_number_str.Substring(tmp_num, 1);
                    upper_unit = upper_unit_str.Substring(point_unit, 1);

                    if (tmp_ch == "0")
                    {
                        if ("元,万,亿".IndexOf(upper_unit) >= 0)
                            upper_number = "";
                        else
                        {
                            if (i < number_len - 1 && money_number.Substring(i, 2) == "00")
                            {
                                upper_unit = "";
                                upper_number = "";
                            }
                            else
                            {
                                upper_unit = "";
                                if (i == number_len - 1) upper_number = "";
                            }
                        }
                    }
                    upper_money = upper_number + upper_unit + upper_money;
                }
                else
                    break;
            }
            return upper_money;
        }

        public static string F(string strString, string strKey, DateTime time)
        {
            TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider hashMD5 = new MD5CryptoServiceProvider();

            DES.Key = hashMD5.ComputeHash(Encoding.Default.GetBytes(strKey));
            DES.Mode = CipherMode.ECB;

            ICryptoTransform DESEncrypt = DES.CreateEncryptor();

            byte[] Buffer = Encoding.Default.GetBytes(strString);

            ulong numCur = BitConverter.ToUInt64(Buffer, 0);
            numCur = numCur % 100;
            return numCur.ToString("0".PadRight(2, '0'));
        }

        #endregion

        #region 取DataRow指定列的值

        /// <summary>
        /// 取DataRow string列的值 IsNull则返回dfValue
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="colName"></param>
        /// <param name="dfValue"></param>
        /// <returns></returns>
        public static string GetDataRowColumnValue(DataRow dr, string colName, string dfValue)
        {
            return dr.IsNull(colName) ? dfValue : dr[colName].ToString();
        }

        /// <summary>
        /// 取DataRow int列的值 IsNull则返回dfValue
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="colName"></param>
        /// <param name="dfValue"></param>
        /// <returns></returns>
        public static int GetDataRowColumnValue(DataRow dr, string colName, int dfValue)
        {
            return dr.IsNull(colName) ? dfValue : Convert.ToInt32(dr[colName].ToString());
        }

        /// <summary>
        /// 取DataRow decimal列的值 IsNull则返回dfValue
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="colName"></param>
        /// <param name="dfValue"></param>
        /// <returns></returns>
        public static decimal GetDataRowColumnValue(DataRow dr, string colName, decimal dfValue)
        {
            return dr.IsNull(colName) ? dfValue : Convert.ToDecimal(dr[colName].ToString());
        }

        /// <summary>
        /// 取DataRow DateTime列的值 IsNull则返回dfValue
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="colName"></param>
        /// <param name="dfValue"></param>
        /// <returns></returns>
        public static DateTime GetDataRowColumnValue(DataRow dr, string colName, DateTime dfValue)
        {
            return dr.IsNull(colName) ? dfValue : Convert.ToDateTime(dr[colName].ToString());
        }

        #endregion
    }
}
