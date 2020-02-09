/*----------------------------------------------------------------
// Copyright (C) 2007��������������޹�˾
//
// �ļ�����ȫ�ַ���
// �ļ�����������
//
// ������ʶ��
//
// �޸ı�ʶ��������
// �޸�������������ʾ�򷽷�
//           
// �޸ı�ʶ��
// �޸�������
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
    /// ϵͳ������
    /// </summary>
    public static class Sysfun
    {
        #region ��������ת��

        /// <summary>
        /// תֵ��DateTime
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
        /// תֵ��Int
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
        /// תֵ��Int
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
        /// תֵ��Long
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
        /// תֵ��Decimal
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
        /// ת��Double
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
        /// ת��String
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
        /// ���ַ�������ת������ɫRGB��ʽ
        /// </summary>
        /// <param name="color3str">��ɫ�ַ���(255,255,255)��255,255,255</param>
        /// <param name="defaultColor">Ĭ����ɫ</param>
        /// <returns>RGB��ɫ</returns>
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
                    MessageBox.Show("��ɫת������! " + e.Message);
                    return defaultColor;
                }
                return color;
            }
        }

        #endregion 

        #region ���ܺ���

        /// <summary>
        /// �ж��ַ����Ƿ�������(ʹ��������ʽ���)
        /// </summary>
        /// <param name="str"></param>
        /// <returns>true/false</returns>
        public static bool IsNumeric(string str)
        {
            if (str == null || str.Length == 0 || str == "-")
                return false;
            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"^[-]?\d*[.]?\d*$");  // ����: 0.74, .74, 10, 10. , 10.1
            return reg.IsMatch(str);
        }

        /// <summary>
        /// �ж��ַ����Ƿ�������(Ҳ,ʹ��������ʽ���, ԭ��:ֻ����,0..9,+-.) Ϊ������ǰ�汾�����˺���
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
            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"^[-]?\d*[.]?\d*$");  // ����: 0.74, .74, 10, 10. , 10.1
            return reg.IsMatch(str);
        }

        /// <summary>
        /// �ж��ַ����Ƿ�������
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
        /// ��ȡ�·ݵ����һ��
        /// </summary>
        /// <param name="DayOfMonth">�·��е�һ��</param>
        /// <returns>���µ����һ��</returns>
        public static DateTime GetLastDateForMonth(DateTime DayOfMonth)
        {
            DateTime DtEnd ; 
            int Dtyear, DtMonth; 
            Dtyear = DayOfMonth.Year; 
            DtMonth = DayOfMonth.Month; 
            int MonthCount = DateTime.DaysInMonth(Dtyear, DtMonth);//��������ж����� 
            DtEnd = Convert.ToDateTime(Dtyear.ToString() + "-" + DtMonth.ToString() + "-" + MonthCount); 
            DtEnd.AddDays(1).AddMilliseconds(-1);
            return DtEnd;
        }


        /// <summary>
        /// ���תȫ��
        /// </summary>
        /// <param name="strChar">ת�����ַ���</param>
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
            ////����ȫ�Ǵ�
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
        /// ȫ��ת���
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

            ////���ذ�Ǵ�
            //return new string(c);
            strChar = strChar.Replace("��", "%");
            strChar = strChar.Replace("��", ".");
            strChar = strChar.Replace("��", ",");
            strChar = strChar.Replace("��", ">");
            strChar = strChar.Replace("��", "<");
            strChar = strChar.Replace("��", "(");
            strChar = strChar.Replace("��", ")");
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
        /// ��ȡ�ַ��������ݿ�ĳ���
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
        /// ����SQL��Where����
        /// </summary>
        /// <param name="originalSql">ԭʼSQL</param>
        /// <param name="condition">���������</param>
        /// <returns>�µ�SQL</returns>
        public static string SQLAddWhere(string originalSql, string condition)
        {
            //�����µ�sqlselect
            string sql = originalSql.ToUpper();
            int li_locBef, li_locLst;

            // �����һ��FROM�������WHERE����
            li_locBef = sql.LastIndexOf("FROM ");
            li_locBef = sql.IndexOf("WHERE ", li_locBef);

            if (li_locBef < 0)  // ��WHERE
            {
                li_locBef = sql.IndexOf("GROUP BY ");
                if (li_locBef < 0)  // ��GROUP
                {
                    li_locBef = sql.IndexOf("ORDER BY ");
                    if (li_locBef < 0)  // ��ORDER
                    {
                        originalSql += " WHERE " + condition + " ";
                    }
                    else  // ��ORDER
                    {
                        originalSql = originalSql.Substring(0, li_locBef - 1) + " WHERE " + condition + originalSql.Substring(li_locBef - 1, originalSql.Length - li_locBef + 1);
                    }
                }
                else  // ��GROUP
                {
                    originalSql = originalSql.Substring(0, li_locBef - 1) + " WHERE " + condition + originalSql.Substring(li_locBef - 1, originalSql.Length - li_locBef + 1);
                }
            }
            else  // ��WHERE
            {
                li_locLst = sql.IndexOf("GROUP BY");
                if (li_locLst >= 0)  // ��GROUP
                {
                    originalSql = originalSql.Substring(0, li_locLst - 1) + " AND (" + condition + ") " + originalSql.Substring(li_locLst - 1, originalSql.Length - li_locLst + 1);
                }
                else
                {
                    li_locLst = sql.IndexOf("ORDER BY");
                    if (li_locLst >= 0)  // ��ORDER
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
        /// ���ز���ֵ�б� 
        /// </summary>
        /// <param name="Para">��ֵ�ַ��� (��;�ָ�)</param>
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
        /// ������ȿ���
        /// </summary>
        /// <param name="SourceObj">Դ����</param>
        /// <param name="TargetObj">Ŀ�����</param>
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
        ///  ת���ɴ�д���
        /// </summary>
        /// <param name="money">���</param>
        /// <returns>��д���</returns>
        public static string UpperMoney(decimal money)
        {
            string upper_number_str = "��Ҽ��������½��ƾ�";
            string upper_unit_str = "ʮ��Ǫ��ʰ��Ǫ��ʰԪ�Ƿ�";
            string upper_number;           //  ��ֵ�Ĵ�д
            string upper_unit;             //  ��λ��д
            string money_number;           //  �������
            string upper_money;            //  ��д���

            int number_len;
            int point_unit;
            string tmp_ch;
            int tmp_num;

            money = System.Decimal.Round(money, 2) * 100.00M;
            if (money == 0.00M)         //    �㷵�ؿ��ַ���
                return "";

            money_number = money.ToString();
            point_unit = money_number.IndexOf('.');
            if (point_unit > 0)
                money_number = money_number.Substring(0, point_unit);
            number_len = money_number.Length;
            upper_money = "��";

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
                        if ("Ԫ,��,��".IndexOf(upper_unit) >= 0)
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

        #region ȡDataRowָ���е�ֵ

        /// <summary>
        /// ȡDataRow string�е�ֵ IsNull�򷵻�dfValue
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
        /// ȡDataRow int�е�ֵ IsNull�򷵻�dfValue
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
        /// ȡDataRow decimal�е�ֵ IsNull�򷵻�dfValue
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
        /// ȡDataRow DateTime�е�ֵ IsNull�򷵻�dfValue
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
