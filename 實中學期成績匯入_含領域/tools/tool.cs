using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;

namespace 實中學期成績匯入_含領域
{
    static public class tool
    {
        static public QueryHelper _Q = new QueryHelper();
        static public List<T>[] Partition2<T>(List<T> list, int size)
        {
            if (list == null)
                throw new ArgumentNullException("list");

            if (size < 1)
                throw new ArgumentOutOfRangeException("size");

            int count = (int)Math.Ceiling(list.Count / (double)size);
            List<T>[] partitions = new List<T>[count];

            int k = 0;
            for (int i = 0; i < partitions.Length; i++)
            {
                partitions[i] = new List<T>(size);
                for (int j = k; j < k + size; j++)
                {
                    if (j >= list.Count)
                        break;
                    partitions[i].Add(list[j]);
                }
                k += size;
            }

            return partitions;
        }
    }
}
