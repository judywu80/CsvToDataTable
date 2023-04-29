using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Data.SqlClient;
using System.Reflection; //新增反射,編譯寫法
using System.Xml.Linq;
using System.Runtime.CompilerServices;
using System.Collections;

namespace CsvToDataTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string sp = ",", cns, sqs; //,path,tname; //int c,r;
        List<Emplo> list; 
        SqlConnection cn; SqlCommand cmd;
        DataTable dt; 

        static int d; //d:Click數; fno:欄位序 () //加static另2新class裡的d才讀的到
        class Emplo: IComparable<Emplo> //內建IComparable<>, 實作CompareTo後才可繼承 //沒實作類別可放靜態區
        {
            public static string sp; //類別內方法ToString需要
            public string 編號 { get; set; } //CsvToDGV:打在這裡DGV就會*出現*欄位
            public string 姓名 { get; set; }
            public int 月薪 { get; set; }
            public DateTime 出生日 { get; set; }
            public string 經歷一 { get; set; }
            public string 經歷二 { get; set; }
            public string 經歷三 { get; set; }
            public static Emplo FromCsv(string line) //建構子 //型別<綠>+方法<黃>(型別+值)
            {
                var split = line.Split(',');
                Emplo emp = new Emplo();
                emp.編號 = split[0];  //右式與上面2行括號內容有關 //填入內容(欄位在上面)
                emp.姓名 = split[1];
                emp.月薪 = int.Parse(split[2]);
                emp.出生日 = Convert.ToDateTime(split[3]);
                emp.經歷一 = split[4];
                emp.經歷二 = split[5];
                emp.經歷三 = split[6];
                return emp;
            }
            public override string ToString() //List to TXT //因標點符號不同,覆寫需ToString (or加new實作化)
            {   //方法ToString,要加return
                return $"{編號}{sp}{姓名}{sp}{月薪}{sp}{出生日.ToString("yyyy/MM/dd")}{sp}{經歷一}{sp}{經歷二}{sp}{經歷三}";//變數內為屬性data&欄位sp
            }   //日期格式轉成無時分  //年用yyyy才是完整年,用yy會省前2位數19
            public int CompareTo(Emplo other) //0個參考:**class裡的會找不到 //Compare'To'+1變數 //實作介面IComparable
            {   
                if (d % 2 == 1) 
                    return 編號.CompareTo(other.編號); 
                else
                    return -編號.CompareTo(other.編號);
            }
        }
        class compareByName: IComparer<Emplo> //是ICompar'er'非comparable //打完繼承會紅字叫我實作化 
        {   //若用編號的寫法"CompareTo",'姓名'會無法辨識,因這是新class無宣告屬性,而上面編號在Emplo類別屬性裡
            public int Compare(Emplo a, Emplo b) //Compare+2個變數
            {
                if (d % 2 == 1)
                    return a.姓名.CompareTo(b.姓名);
                else
                    return -a.姓名.CompareTo(b.姓名);
            }
        }
        private void Form1_Load(object sender, EventArgs e)  //原csv to listBox
        {
            cns = "Data Source=.\\SQL2019; Database= Emplo;" + //**改主機名**&資料庫    .\\sql2019
            //      "User ID= sa; PWD= **";       //**改密碼**
                  "Trusted_Connection= True";    //改信任方式登入
            sqs = "Select * From Employee";    //改資料表名 tname = "Emplyee";

            string[] ary = { "逗號", "分號", "空白", "TAB", "冒號" }; //或寫在元件>屬性>資料item的集合裡
            comboBox1.Items.AddRange(ary);  //AddRange複數
            /*listBox1.Text = "";          //原csv to listBox
            listBox1.Items.AddRange(File.ReadAllLines("..\\..\\B01.csv",Encoding.Default)); //將陣列加到listBox*/
        }
        private void btn1_Click(object sender, EventArgs e) //Csv To DGV //原listBox to dGV
        {   //CsvToDGV:泛型資料內容(不含欄位,欄位寫在靜態類別DGV會自已顯現)
            dataGridView1.DataSource = null; //or資料會一直存在,按其他btn無法顯現

            list = new List<Emplo>(); //實作化在btn裡免恆顯示
            list = File.ReadAllLines("..\\..\\B01.csv", Encoding.Default).Skip(1).Select(x => Emplo.FromCsv(x)).ToList(); //轉泛型-實作化
            //**藉由類別方法，將讀取的資料實做泛形 list  //讀Select用法
            dataGridView1.DataSource = list; //**用list泛型btn1會無法排序
            
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
            dataGridView1.EnableHeadersVisualStyles = false;
            columnHeaderStyle.BackColor = Color.FromArgb(204, 255, 204);
            columnHeaderStyle.ForeColor = Color.FromArgb(0, 0, 255);                  //字體色 藍色;
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            DataGridViewCellStyle rowHeaderStyle = new DataGridViewCellStyle();
            dataGridView1.EnableHeadersVisualStyles = false;
            rowHeaderStyle.BackColor = Color.FromArgb(204, 255, 204);
            dataGridView1.RowHeadersDefaultCellStyle = rowHeaderStyle;
            dataGridView1.Font = new Font("Arial", 11, FontStyle.Regular);            //字體大小含欄位
            dataGridView1.ForeColor = Color.FromArgb(25, 25, 112);
            dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(255, 222, 222);                 //Cells背景色 淺橘色
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(171, 255, 255);  //偶數列背景色 淺藍色
            dataGridView1.AutoResizeColumns();    //欄寬自動調整

            /*dataGridView1.Rows.Clear();  //原listBox to dGV
            string[] lst = listBox1.Items[0].ToString().Split(','); //取listBox1的第一列(Items[0]) //"第一列"用逗點分隔存成陣列
            dataGridView1.ColumnCount=c= lst.Count();  //陣列後加Count方法() //7欄 //可用2個等號表示
            dataGridView1.RowCount=r = listBox1.Items.Count;  //r=98列 //等號左式放未知,右式放已知
            for (int j=0; j<c; j++)  //宣告int j, j=0因陣列第一個位置=0
            {
                dataGridView1.Columns[j].Name=lst[j];  //Columns"[j]"只有欄資料要填
            }
            for (int i=0; i<r-1; i++)  //資料少欄位第一列=98-1
            {
                lst = listBox1.Items[i + 1].ToString().Split(','); //"每列資料"用逗點分隔,存到成陣列
                for (int j=0; j<lst.Length; j++)  //可用.Length=.Count()=7
                { 
                    if (j == 2)
                    { dataGridView1[j, i].Value = int.Parse(lst[j]); }  //日期也要轉,Convert.ToDateTime(lst[j]).ToString("yyyy/MM/dd")
                    else if(j==3)
                    { dataGridView1[j,i].Value=Convert.ToDateTime(lst[j]).ToString("yyyy/MM/dd"); } //不然日期照字串排,非數字,如12月大於3月 //指定年月日不然會有完整詳細時間(月日前面會補0)
                    else
                    { dataGridView1[j, i].Value = lst[j]; }
                }
            }*/
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //上面Split那段,多設1參數sp
        {
            int x = comboBox1.SelectedIndex; //改用index位置,免打中文
            switch (x)
            {
                case 0: sp = ","; break;
                case 1: sp = ";"; break;
                case 2: sp = " "; break;
                case 3: sp = "\t"; break; //用tab打出來的空格比用\t來的小
                case 4: sp = ":"; break;
                default: sp = ","; break; //不選擇就是逗號,因靜態變數已宣告=","
            }
            /*if (comboBox1.SelectedItem.ToString() == "逗號") sp = ","; //原button1參數sp會自己代入sp (從dGV to txt)
            if (comboBox1.SelectedItem.ToString() == "分號") sp = ";";
            if (comboBox1.SelectedItem.ToString() == "空白") sp = " ";
            if (comboBox1.SelectedItem.ToString() == "TAB") sp = "\t";
            if (comboBox1.SelectedItem.ToString() == "冒號") sp = ":"; */
        }
        private void button3_Click(object sender, EventArgs e)
        {   //dt=null;    //似無影響
            dataGridView1.DataSource = null;

            cn = new SqlConnection(cns);
            cmd = new SqlCommand(sqs, cn);
            cn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            dt = new DataTable();
            dt.Load(dr);    //載入SqlDataReader的資料  //原dr(from db)有資料,讀db上dt的資料
            //dr.Close(); cmd.Dispose(); //似無變化
            cn.Close(); //**

            dataGridView1.DataSource = dt; //Table會也自動排,數字&日期也排對 //如果使用DataView還可以有多欄位排序，及filte功能
            dataGridView1.Font = new Font("Arial", 11, FontStyle.Regular); //字體大小含欄位
            dataGridView1.AutoResizeColumns(); //此行放最後,才會依字大小調整 //欄寬自動調整 
        }
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {   //S.Sort((x, y) => x.Score.CompareTo(y.Score)); //多設btn只有得分排序
            int indx = e.ColumnIndex; //滑鼠欄位c位置
            int z; //+1(ASC)or-1(DESC) //用三元運算子前面還可順便宣告
            d += 1; ////d滑鼠按的次數，作遞增或遞減控制 //**當使用者按了欄位列的某一個行銷e.ColumnIndex，d就會加1。
            //int z = d % 2 == 1 ? 1 : -1;
            if (d % 2 == 1) { z = 1; } //沒加中括號會有問題
            else { z = -1; } //沒加中括號會有問題
            //list.Sort(); 預設有4種(多載)
            if (indx == 0)    
                list.Sort();  //1.(編號)使用IComparable<ItemData>
            else if (indx == 1)  
                list.Sort(new compareByName());  //2.自定義class compareByName: IComparer<Emplo>
            else if (indx == 2) 
                list.Sort(delegate (Emplo a, Emplo b){return z * a.月薪.CompareTo(b.月薪); });
                //3.委派DeleGate表達式(?):似把方法2.類別內容寫在這裡 //Lambda如下更簡潔
                //**呼叫方法是由方法自行處理，委派是將如何處理事情的方法委派其他方法處理。
            else if (indx == 3) 
                list.Sort((x, y) => { return z * x.出生日.CompareTo(y.出生日); }); //4.Lambda表達式
            else if (indx == 4)
                list.Sort((x, y) => { return z * x.經歷一.CompareTo(y.經歷一); });  
                //5.下次再用由欄位名稱"反射"欄位屬性的List.OrderBy(),程式更精簡: ListIt = ListIt.OrderBy(x => x.GetType().GetProperties()[vc].GetValue(x, null)).ToList();
            else if (indx == 5)
                list.Sort((x, y) => { return z * x.經歷二.CompareTo(y.經歷二); });
            else if (indx == 6)
                list.Sort((x, y) => { return z * x.經歷三.CompareTo(y.經歷三); });
            //**bindingList的bs重新綁定(類bn的bs), 若用反射寫法(仍然需重新綁定)不需逐欄判斷**

            var bL = new BindingList<Emplo>(list); //list加入 
            var sou=new BindingSource(bL, null);
            dataGridView1.DataSource = sou;
        }
        private void button1_Click(object sender, EventArgs e)  //dGV(List) to txt
        {   //按兩次會跳出2個相同檔案
            Emplo.sp = sp; //宣告使用類別.欄位sp=sp標點符號 //**因類別Emplo有定義 sp (field)，ToString()會用到，需要指派。
            using (StreamWriter sw = new StreamWriter("..\\..\\test.txt")) //寫入時會自己新增txt檔不用先手動新增
            {   //欄位寫入方法不變,無法用泛型寫入txt,因(非dGV)是txt格式=泛形資料無欄位，要用Emplo的屬性名稱，本例直用dgv的欄位名稱
                for (int i = 0; i < dataGridView1.ColumnCount; i++) //因j逗號要變,故拆開先填i=0單字"編號",非上面j直接一列
                {
                    sw.Write(dataGridView1.Columns[i].HeaderText); //只有Write,不是WriteLine(會變每欄換行)
                    if (i != dataGridView1.Columns.Count - 1) //不是最後一項,後面都加符號
                    { sw.Write(sp); }
                }
                sw.WriteLine();
                //內容寫入方法有變,用泛型. //用不同的泛型名data因data是泛形list的一筆，並不等於list //因加入標點符號sp,顯現用ToString列印
                foreach (var data in list) //list= new List<Emplo>();
                {
                    sw.WriteLine(data.ToString(), Encoding.Default); //list透過str ToString & WriteLine寫入TXT檔 //移除,Encoding.Default中文會出現亂碼
                }
                /*for (int i=0;i<dataGridView1.RowCount-1;i++) //保留原dGV to txt內容寫法對照
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)  //這裡是c,二維
                    {
                        sw.Write(dataGridView1.Rows[i].Cells[j].Value.ToString()); //用法.Cell[].Value.ToString
                        if (j != dataGridView1.Columns.Count - 1)
                        { sw.Write(sp); }
                    }
                    if (i != dataGridView1.Rows.Count - 2)     //**不使文字檔末列空行(末列不需換行)
                        sw.WriteLine();      //沒這行,內容完全不會換行
                }*/
            }
            Process.Start("..\\..\\test.txt"); //跳出寫入檔案 //不需用FileStream @= File.Create @2
        }
        private void button2_Click(object sender, EventArgs e) //存入db需Adapter, 先用adapter開檔案（說明於scb）
        {   //寫入實體Table之前，請採清除原資料表。//dt=null;
            if (dataGridView1.DataSource == null) 
            { 
                MessageBox.Show("請先按「執行轉換」讀csv檔");
                return; //跳回未按此按鈕前,程式仍在執行
            }
            dataGridView1.DataSource=null;
            
            cn = new SqlConnection(cns);
            cmd= new SqlCommand(); //先不指定sqs內容,因有2個(1先刪資1讀全)如下
            //下cmd del指令(可用解構子),or adapt.Update(dt)違反PRIMARY KEY條件。
            cn.Open(); //有adapter還要開啟
            cmd.Connection = cn;
            //cmd.CommandType = CommandType.Text; //沒寫也可 //預設值即為TEXT //補充:CommandType是SqlCommand对象的一个属性，用于指定执行动作的形式，它告诉.net接下来执行的是一个文本(text)、存储过程(StoredProcedure)还是表名称(TableDirect).
            cmd.CommandText = "Truncate Table Employee"; //delete from t(sql語法) //truncate比delete from快(drop是完全刪).
            cmd.ExecuteNonQuery(); //刪除表可用解構子using，解構子（最常用在資源管理），能在Class在一開始透過建構子與資料庫連線，取得內容，再透過解構子，在Class結束前關閉連線，或是在解構子中釋放記憶體變數。
            cn.Close(); //**加此行,or CRUD無法連續作業

            SqlDataAdapter adapt = new SqlDataAdapter(sqs, cn); //若有2個cmd sql str語法(1刪除.1查詢全),會找不到(第一個)資料表0.
            //**使用cmd通常是下命令去取資料或直接異動資料，本案例是採用DataSet，必須有adapt，又採用adapt.Update()，就必須有 SqlCommandBuilder scb
            SqlCommandBuilder scb = new SqlCommandBuilder(adapt); //直接用 Update 而非使用 UpdateCommand 時用 //*需要此行*,or最下行adapt.Update(dt)會'當傳遞擁有新資料列的 DataRow 集合時，Update 需要有效的 InsertCommand。'
            DataSet ds = new DataSet();
            adapt.Fill(ds, "Employee");  //置入Dataset ds 
            dt = ds.Tables[0];  //0=第一個table,在DataSet的第1個DataTable //並取 DataTable tb
            
            //Get all the properties //下句使用反射編譯寫法,實作化Emplo並找到所有屬性
            PropertyInfo[] prop = typeof(Emplo).GetProperties(BindingFlags.Instance | BindingFlags.Public);//自動填入將NonPublic改為Public //typeof(x)等同x.GetType()
            //陣列 //BindingFlags 列舉: 指定控制繫結的旗標和由反映執行的成員和類型搜尋方式。反射查找方法,默认查找public、instance(實作化)内容
            foreach (Emplo p in list) //巡覽二維陣列in泛型 //原字串寫入時與override ToString()衝突=被截斷,因1字串存入txt&另1存入db,已修復
            {
                var value = new object[prop.Length];
                for (int i = 0; i < prop.Length; i++)
                    value[i] = prop[i].GetValue(p, null); //預寫dt改p //inserting property values to datatable rows
                dt.Rows.Add(value); //將(97筆row)資料逐筆新增到 DataTable 中
            }

            adapt.Update(dt); //更新C(rud) //改變在view,沒真正寫進table //原未下cmd指令時,'當傳遞擁有新資料列的 DataRow 集合時，Update 需要有效的 InsertCommand。'
            dataGridView1.DataSource = dt; //先讀 
        }
    }
}
