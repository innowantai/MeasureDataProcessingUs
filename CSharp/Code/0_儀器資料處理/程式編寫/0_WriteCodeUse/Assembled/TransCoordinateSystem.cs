using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Assembled
{
    public class TransCoordinateSystem
    {


        private List<double> Parameter_y = new List<double>() {2800000,2750000,2700000,
                                                              2650000,2600000,2550000,2500000,
                                                              2450000,2400000    };
        private List<double> Paramtert_x = new List<double>() { 90000, 170000, 250000, 330000, 410000 };

        public TransCoordinateSystem()
        {


        }



        public void Main_TWD97toTWD67()
        {
            double input_x = 165557.000;
            double input_y = 2563222.000;
        }




    }
}
