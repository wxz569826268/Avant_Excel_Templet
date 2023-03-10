using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplet
{
    public static class BaseMethod
    {
        /// <summary>
        /// 数据表头
        /// </summary>
        public struct DataRows
        {
            /// <summary>
            /// 列名
            /// </summary>
            public string[] Columns;
            /// <summary>
            /// 列数据类型Ex. typeof(string)
            /// </summary>
            public Object[] type;
        }

        /// <summary>
        /// 创建数据源
        /// </summary>
        /// <param name="Columns">列名及类型</param>
        /// <param name="LI">数据集</param>
        /// <returns></returns>
        public static DataTable BuidData(DataRows Columns, List<Object[]> LI)
        {

            DataTable data = new DataTable();
            for (int i = 0; i < Columns.Columns.Length; i++)
            {
                data.Columns.Add(Columns.Columns[i], (Type)Columns.type[i]);
            }
            DataRow DR = data.NewRow();
            for (int i = 0; i < LI.Count; i++)
            {
                DR = data.NewRow();
                for (int j = 0; j < LI[i].Length; j++)
                {
                    DR[j] = LI[i][j];
                }
                data.Rows.Add(DR);
            }
            return data;
        }
    }
    }
