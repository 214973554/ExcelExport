using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entity
{
    public class BanJiTongXunLuHead : BaseHead
    {
        public string BanJi { get; set; }

        public string BanZhuRen { get; set; }

        public string BanZhang { get; set; }
    }

    public class BanJiTongXunLuListRow : BaseListRow
    {
        public int XueHao { get; set; }

        public string XingMing { get; set; }

        public string XingBie { get; set; }

        public string LianXiDianHua { get; set; }

        public string QQ { get; set; }

        public string WeiXin { get; set; }

        public string ShiFouZhuSu { get; set; }

        public string JiaZhangXingMing { get; set; }

        public string JiaZhangDianHua { get; set; }

        public string JiaTingZhuZhi { get; set; }


    }
}
