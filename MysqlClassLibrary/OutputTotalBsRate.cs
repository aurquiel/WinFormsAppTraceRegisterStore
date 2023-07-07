using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MysqlClassLibrary
{
    public class OutputTotalBsRate
    {
        private List<OutputTotalBsRate> listValues = new List<OutputTotalBsRate>();
        private decimal rate { get; set; }
        private decimal totalBs { get; set; }
        private int processedStatus { get; set; }

        public decimal Rate
        {
            get
            {
                return rate;
            }

            set
            {
                rate = value;
            }
        }

        public decimal TotalBs
        {
            get
            {
                return totalBs;
            }

            set
            {
                totalBs = value;
            }
        }

        public int ProcessedStatus
        {
            get
            {
                return processedStatus;
            }

            set
            {
                processedStatus = value;
            }
        }

        public List<OutputTotalBsRate> ListValues
        {
            get
            {
                return listValues;
            }

            set
            {
                listValues = value;
            }
        }
    }
}
