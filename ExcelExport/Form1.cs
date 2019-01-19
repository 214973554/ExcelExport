using Entity;
using Framework.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = Path.Combine(Application.StartupPath, "Template", "模板.xlsx");
            var data = GetDataSource();
            string savePath = Path.Combine(Application.StartupPath, "Template", string.Format("{0}.xlsx", Guid.NewGuid()));

            ExcelHelper.ExportExcel(path, savePath, data);

            MessageBox.Show("生成成功。");
        }

        private List<SheetDataSource<BaseHead, BaseListRow>> GetDataSource()
        {
            var chengJiTongJiBiaoDataSource = GetChengJiTongJiBiaoDataSource();
            var banJiTongXunLuDataSource = GetBanJiTongXunLuDataSource();
            var banJiTongXunLuListDataSource = GetBanJiTongXunLuListDataSource();

            var list = new List<SheetDataSource<BaseHead, BaseListRow>>();
            list.Add(chengJiTongJiBiaoDataSource);
            list.Add(banJiTongXunLuDataSource);
            list.Add(banJiTongXunLuListDataSource);

            return list;
        }

        private SheetDataSource<BaseHead, BaseListRow> GetChengJiTongJiBiaoDataSource()
        {
            var ds = new DataSource<BaseHead, BaseListRow>();
            
            Head head = new Head();
            head.School = "清华大学";
            head.Class = "大一";
            head.Semester = "第二学期";
            head.TeacherInCharge = "朱自清";
            head.OpenDate = null;

            ds.FieldData = head;

            ds.ListData = new List<BaseListRow>();

            ds.ListData.Add(new Grade()
            {
                XueHao = 1,
                XingMing = "张三",
                YuWen = 10,
                ShuXue = 20,
                YingYu = 30,
                SiXiangPingDe = 40,
                LiShi = 50,
                WuLi = 60,
                HuaXue = 70,
                TiYu = 80,
                ZongFen = 999
            });

            ds.ListData.Add(new Grade()
            {
                XueHao = 2,
                XingMing = "李四",
                YuWen = 100,
                ShuXue = 200,
                YingYu = 300,
                SiXiangPingDe = 400,
                LiShi = 500,
                WuLi = 600,
                HuaXue = 700,
                TiYu = 800,
                ZongFen = 9990
            });
            


            var chengJiTongJiBiaoDataSource = new SheetDataSource<BaseHead, BaseListRow>();
            chengJiTongJiBiaoDataSource.SheetName = "成绩统计表";
            chengJiTongJiBiaoDataSource.DataSource = ds;

            return chengJiTongJiBiaoDataSource;
        }

        private SheetDataSource<BaseHead, BaseListRow> GetBanJiTongXunLuDataSource()
        {
            var  ds = new DataSource<BaseHead, BaseListRow>();

            var head = new BanJiTongXunLuHead();
            head.BanJi = "初三（8）班";
            head.BanZhuRen = "张艳";
            head.BanZhang = "刘洋";

            ds.FieldData = head;

            ds.ListData = new List<BaseListRow>();
            ds.ListData.Add(new BanJiTongXunLuListRow()
            {
                XueHao = 1,
                XingMing = "张三",
                XingBie = "男",
                LianXiDianHua = "123****6666",
                QQ = "123**666",
                WeiXin = "1234**66665",
                ShiFouZhuSu = "是",
                JiaZhangXingMing = "张三爸",
                JiaZhangDianHua = "123****123",
                JiaTingZhuZhi = "中国上海"
            });

            ds.ListData.Add(new BanJiTongXunLuListRow()
            {
                XueHao = 2,
                XingMing = "张三2",
                XingBie = "男2",
                LianXiDianHua = "123****6666-2",
                QQ = "123**666-2",
                WeiXin = "1234**66665-2",
                ShiFouZhuSu = "是-2",
                JiaZhangXingMing = "张三爸-2",
                JiaZhangDianHua = "123****123-2",
                JiaTingZhuZhi = "中国上海-2"
            });


            var  banJiTongXunLuDataSource = new SheetDataSource<BaseHead, BaseListRow>();
            banJiTongXunLuDataSource.SheetName = "班级通讯录";
            banJiTongXunLuDataSource.DataSource = ds;

            return banJiTongXunLuDataSource;
        }

        private SheetDataSource<BaseHead, BaseListRow> GetBanJiTongXunLuListDataSource()
        {
            var ds = new DataSource<BaseHead, BaseListRow>();

            ds.ListData = new List<BaseListRow>();
            ds.ListData.Add(new BanJiTongXunLuListRow()
            {
                XueHao = 3,
                XingMing = "张三",
                XingBie = "男",
                LianXiDianHua = "123****6666",
                QQ = "123**666",
                WeiXin = "1234**66665",
                ShiFouZhuSu = "是",
                JiaZhangXingMing = "张三爸",
                JiaZhangDianHua = "123****123",
                JiaTingZhuZhi = "中国上海"
            });

            ds.ListData.Add(new BanJiTongXunLuListRow()
            {
                XueHao = 4,
                XingMing = "张三2",
                XingBie = "男2",
                LianXiDianHua = "123****6666-2",
                QQ = "123**666-2",
                WeiXin = "1234**66665-2",
                ShiFouZhuSu = "是-2",
                JiaZhangXingMing = "张三爸-2",
                JiaZhangDianHua = "123****123-2",
                JiaTingZhuZhi = "中国上海-2"
            });


            var banJiTongXunLuListDataSource = new SheetDataSource<BaseHead, BaseListRow>();
            banJiTongXunLuListDataSource.SheetName = "班级通讯录-List";
            banJiTongXunLuListDataSource.DataSource = ds;

            return banJiTongXunLuListDataSource;
        }
    }
}
