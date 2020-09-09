using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
public static class ExFuncsByLbl
{
    [ExcelFunction(Description = "农历转换")]
    public static DateTime ChangeDate([ExcelArgument(Name ="日期")] DateTime date, [ExcelArgument(Name = "类型",Description ="0为农历转阳历,1为阳历转农历")] int type=0)
    {
        
        int[] config = new int[]{0x04bd8, 0x04ae0, 0x0a570, 0x054d5, 0x0d260, 0x0d950, 0x16554, 0x056a0, 0x09ad0, 0x055d2,
0x04ae0, 0x0a5b6, 0x0a4d0, 0x0d250, 0x1d255, 0x0b540, 0x0d6a0, 0x0ada2, 0x095b0, 0x14977,
0x04970, 0x0a4b0, 0x0b4b5, 0x06a50, 0x06d40, 0x1ab54, 0x02b60, 0x09570, 0x052f2, 0x04970,
0x06566, 0x0d4a0, 0x0ea50, 0x06e95, 0x05ad0, 0x02b60, 0x186e3, 0x092e0, 0x1c8d7, 0x0c950,
0x0d4a0, 0x1d8a6, 0x0b550, 0x056a0, 0x1a5b4, 0x025d0, 0x092d0, 0x0d2b2, 0x0a950, 0x0b557,
0x06ca0, 0x0b550, 0x15355, 0x04da0, 0x0a5d0, 0x14573, 0x052d0, 0x0a9a8, 0x0e950, 0x06aa0,
0x0aea6, 0x0ab50, 0x04b60, 0x0aae4, 0x0a570, 0x05260, 0x0f263, 0x0d950, 0x05b57, 0x056a0,
0x096d0, 0x04dd5, 0x04ad0, 0x0a4d0, 0x0d4d4, 0x0d250, 0x0d558, 0x0b540, 0x0b5a0, 0x195a6,
0x095b0, 0x049b0, 0x0a974, 0x0a4b0, 0x0b27a, 0x06a50, 0x06d40, 0x0af46, 0x0ab60, 0x09570,
0x04af5, 0x04970, 0x064b0, 0x074a3, 0x0ea50, 0x06b58, 0x055c0, 0x0ab60, 0x096d5, 0x092e0,
0x0c960, 0x0d954, 0x0d4a0, 0x0da50, 0x07552, 0x056a0, 0x0abb7, 0x025d0, 0x092d0, 0x0cab5,
0x0a950, 0x0b4a0, 0x0baa4, 0x0ad50, 0x055d9, 0x04ba0, 0x0a5b0, 0x15176, 0x052b0, 0x0a930,
0x07954, 0x06aa0, 0x0ad50, 0x05b52, 0x04b60, 0x0a6e6, 0x0a4e0, 0x0d260, 0x0ea65, 0x0d530,
0x05aa0, 0x076a3, 0x096d0, 0x04bd7, 0x04ad0, 0x0a4d0, 0x1d0b6, 0x0d250, 0x0d520, 0x0dd45,
0x0b5a0, 0x056d0, 0x055b2, 0x049b0, 0x0a577, 0x0a4b0, 0x0aa50, 0x1b255, 0x06d20, 0x0ada0 };
        DateTime STARTDATE = new DateTime(1900, 1, 31);
        
        Dictionary<int, KeyValuePair<int,int>[]> configdic = new Dictionary<int, KeyValuePair<int, int>[]>();
        for(int j=0;j<config.Length;j++)
        {
            string str1 = Convert.ToString(config[j],2).PadLeft(20,'0');
            
            int l = 12 + (str1.Substring(16, 4) == "0000" ? 0 : 1);
            
            List<KeyValuePair<int, int>> ls = new List<KeyValuePair<int, int>>();

            double[] monthcnt = new double[l];
            for(int i = 0; i < 12; i++)
            {
                KeyValuePair<int, int> tmp = new KeyValuePair<int, int>(i + 1, 29 + (str1.Substring(4 + i, 1) == "0" ? 0 : 1));
                ls.Add(tmp);
            }
            if (l == 13)
            {
                KeyValuePair<int, int> tmp = new KeyValuePair<int, int>(Convert.ToInt32(str1.Substring(16, 4),2), 29 + (str1.Substring(0, 4) == "0000" ? 0 : 1));
                ls.Add(tmp);
            }
            
            configdic.Add(1900+j, ls.OrderBy(h=>h.Key).ToArray());
        }
        Dictionary<int,int[]> c = new Dictionary<int,int[]>();
        Dictionary<Int64, int> pvtc = new Dictionary<Int64, int>();
        
        //sumday,year,month,day
        int sumday = -1;
        foreach (KeyValuePair<int, KeyValuePair<int, int>[]> kp in configdic)
        {
            int y = kp.Key;
            foreach(KeyValuePair<int,int> hh in kp.Value)
            {
                int m = hh.Key;
                for(int i = 0; i < hh.Value; i++)
                {
                    int d = i + 1;
                    sumday = sumday + 1;
                    c.Add(sumday, new int[] { y, m, d });
                }
            }
   
        }
        foreach(KeyValuePair<int,int[]> kp in c)
        {
            Int64 key = kp.Value[0]*10000+kp.Value[1]*100+kp.Value[2] ;
            if (!pvtc.ContainsKey(key))
            {
                pvtc.Add(key, kp.Key);
            }
            
        }
        

        if (type == 0)
        {
            int key = date.Year * 10000 + date.Month * 100 + date.Day;
            int deltadays = pvtc[key];
            return STARTDATE.AddDays(deltadays);

        }
        else
        {
            int deltadays = (date - STARTDATE).Days;
            int[] hh = c[deltadays];
            return new DateTime(hh[0], hh[1], hh[2]);
        }


        
    }
}