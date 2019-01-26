using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace _2015ck_1
{
    public partial class Form1 : Form
    {
        internal string walkDir = "";
        int listCount = 0;
        int checkTab = 0; //탭이 1번인지 2번인지 구분 (1번 탭 -> checkTab == 0)

        //리스트뷰1에 필요한 변수
        int liNo1 = 1; //listView1 fNo1 -> 파일 숫자
        string liNa1; //listView1 fName1 -> 파일 이름+확장자
        string liSi1; //listView1 fSize1 -> 파일 크기
        string liPa1; //listView1 fPath1 -> 파일 절대경로
        string liCr1, liAc1; //listView1 fCreate1, fAccess1 -> 파일 생성 시간, 마지막 접근 시간
        string liHa1 = ""; //listView1 fHash1 -> 파일 해쉬(SHA1)
        string lower; //현재 파일확장자 ToLower()
        string liReli1;
        //리스트뷰2에 필요한 변수
        int liNo2 = 1; //listView2 fNo2 -> 파일 숫자
        string liNa2; //listView2 fName2 -> 파일 이름
        string liPa2; //listView2 fPath2 -> 파일 절대경로
        string lipEx2, lirEx2; //listView2 fpExt2, frExt2 -> 현재 파일확장자, 실제 파일확장자
        string liSi2; //listView2 fSig2 -> 파일 시그니처
        string liHa2; //listView2 fHash2 -> 파일 해쉬(SHA1)

        int count, count1; //count -> ' '으로 split된 시그니처 개수
        int x; //for문을 위해 필요한 변수
        int chColor = 0;
        int getExtcount = 0;
        int sigcount = 0;

        public Form1()
        {
            InitializeComponent();
            connect();
            listView1.AllowDrop = true;
            listView2.AllowDrop = true;
            button1.Enabled = false;
            button3.Enabled = false;
            label1.Text="파일 개수 : 0";
            label2.Text = "e-hacking.org";
            label2.Show();
            listCount = listView1.Items.Count;
            if(listCount <= 0)
                button2.Enabled = false;
            //System.DateTime date1 = new System.DateTime(2015, 12, 2, 22, 0, 0);
        }

        int fCount = 0; //검색하고자 하는 디렉터리의 전체 파일 개수
        int z; //시그니처와 확장자명 매치시킨 개수

        public class share //시그니처
        {
            public static string url = @"C:\\Users\\S0NG2\\Desktop\\Capstone,CK-1\\2015ck-1\\Signature_list.txt";
            public static string[] Save_Data = new string[100];
        }

        private void connect()
        {
            try
            {
                string textValue = System.IO.File.ReadAllText(share.url);
                string[] split = textValue.Split('\n');
                z = split.Length;
            }
            catch
            {
                MessageBox.Show("파일이 있는지 확인하고 다시 실행시켜주세요");
                button4.Enabled = false;
            }
        }

/*
        private void connect() //url에 연결하여 시그니처와 확장자 매치한 것 한 줄씩 받아오는 함수
        {
            try
            {
                HttpWebRequest hwr = (HttpWebRequest)WebRequest.Create(share.url);
                HttpWebResponse hwsp = (HttpWebResponse)hwr.GetResponse();
                StreamReader sr = new StreamReader(hwsp.GetResponseStream(), Encoding.Default);

                z = 0;

                while (sr.Peek() > -1) // 서버에 있는 확장자와 시그니처 매치해둔것을 한줄씩 받아와 Save_Data에 저장
                {
                    string Rev_Data = sr.ReadLine();
                    share.Save_Data[z] = Rev_Data;
                    z++;
                }
            }
            catch
            {
                MessageBox.Show("네트워크를 연결을 확인하고 다시 실행시켜주세요");
                button4.Enabled = false;
            }
        }
*/

        private string GetFileSize(double byteCount) //파일의 사이즈를 구할 때 사용하는 함수
        {
            string size = "0 Bytes";
            if (byteCount >= 1073741824.0)
                size = String.Format("{0:##.##}", byteCount / 1073741824.0) + " GB";
            else if (byteCount >= 1048576.0)
                size = String.Format("{0:##.##}", byteCount / 1048576.0) + " MB";
            else if (byteCount >= 1024.0)
                size = String.Format("{0:##.##}", byteCount / 1024.0) + " KB";
            else if (byteCount > 0 && byteCount < 1024.0)
                size = byteCount.ToString() + " Bytes";

            return size;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkTab == 0) //2번 탭을 선택했다면
                checkTab = 1; //checkTab을 1로
            else //1번 탭을 선택했다면
                checkTab = 0; //checkTab을 0으로

            if (checkTab == 0) //첫번째 탭일 때 전체 파일의 개수를 얻어옴
            {
                listCount = listView1.Items.Count;
                label1.Text = "파일 개수 : " + listCount;
                label1.Show();
            }
            else //두번째 탭일 때 시그니처와 확장자가 다른 파일의 개수를 얻어옴
            {
                listCount = listView2.Items.Count;
                label1.Text = "파일 개수 : " + listCount;
                label1.Show();
            }
        }

        private void button1_Click(object sender, EventArgs e) //start버튼을 눌렀을 때
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            listView1.Items.Clear();
            listView2.Items.Clear();
            textBox1.Text = null;
            DateTime localDate1 = DateTime.Now;
            textBox1.AppendText("탐색 시작, " + localDate1 + "\n");
            liNo1 = 1;
            liNo2 = 1;
            DirectoryInfo pDir = new DirectoryInfo(walkDir);
            walkDirectory(pDir); //*****지정된 경로부터 최하위 폴더까지의 모든 파일을 얻어옴
            if (checkTab == 0)
            {
                listCount = listView1.Items.Count;
                label1.Text = "파일 개수 : " + listCount;
                label1.Show();
            }
            else
            {
                listCount = listView2.Items.Count;
                label1.Text = "파일 개수 : " + listCount;
                label1.Show();
            }
            button1.Enabled = true;
            if (listCount > 0)
                button2.Enabled = true;
            else
                button2.Enabled = false;
            button3.Enabled = true;
            button4.Enabled = true;
            DateTime localDate2 = DateTime.Now;
            TimeSpan t3 = localDate2.Subtract(localDate1);
            string time = t3.ToString().Substring(0, 11);
            //MessageBox.Show(time);
            textBox1.AppendText("탐색 종료, " + localDate2 + "\n");
            MessageBox.Show("탐색이 종료되었습니다.\n총 스캔시간 : " + time, "탐색 시간");
        }

        private void cpFile(FileInfo file)
        { //파일의 확장자와 시그니처를 비교
            sigcount = 0;
            chColor = 0;
            x = 0;
            string fullFile = file.FullName;
            string pExtension = "";
            //MessageBox.Show(file.Extension);
            if (file.Extension == "") //파일의 확장자가 없다면
                pExtension = "";
            else
                pExtension = file.Extension.Substring(1);
                //pExtension = "  ";
            lower = pExtension.ToLower();
            liPa1 = fullFile;
            liNa1 = fullFile.Substring(fullFile.LastIndexOf('\\') + 1);
            liCr1 = file.CreationTime.ToString();
            liAc1 = file.LastAccessTime.ToString();
            liSi1 = GetFileSize(file.Length);
            liReli1 = "Lower";
            //liHa1 = ComputeSHA1Hash(fullFile);
            liNa2 = liNa1;
            liPa2 = liPa1;
            //liHa2 = liHa1;

            if (file.Length > 200000000) // 파일이 너무 크면 DeepPink색으로 마크업하고 리스트뷰에 그냥 띄움 
            { //시그니처랑 확장자 비교하지 않는다.
                //MessageBox.Show("파일이 너무 커");
                ListViewItem item2 = new ListViewItem("" + liNo1);
                item2.SubItems.Add(liNa1);
                item2.SubItems.Add(liSi1);
                item2.SubItems.Add(liPa1);
                item2.SubItems.Add(liCr1);
                item2.SubItems.Add(liAc1);
                item2.SubItems.Add(liHa1);
                item2.SubItems.Add(liReli1);
                item2.BackColor = Color.DeepPink;
                listView1.Items.Add(item2);
                liNo1++;
            }
            else //파일의 크기가 200MB 이하일 때
            {
                if (/*file.Length < 200000000 && */file.Length > 20)
                { //파일의 크기가 20byte가 넘을 때 
                    for (x = 0; x < z; x++)
                    {
                        BinaryReader br = new BinaryReader(File.OpenRead(fullFile)); //파일을 열어 바이너리로 읽음
                        br.BaseStream.Position = 0;
                        count = 0; count1 = 0; getExtcount = 0;
                        liSi2 = "";
                        string[] Ex_Def = share.Save_Data[x].Split(' ');
                        foreach (string ct in Ex_Def)
                        { //하나에 시그니처에 대해 바이트 개수를 얻어옴 -> (11 22 33 44 = qp) 하나의 시그니처 개수는 4개
                            if (ct == "=")
                                break;
                            count++;
                        }
                        string[] lotsExt = Ex_Def[count + 1].Split(','); //확장자 저장. 확장자가 여러개일 때 ,으로 나눠 저장
                        string comExt = Ex_Def[count + 1];

                        foreach (string ext in lotsExt)
                        { //하나의 시그니처에 확장자가 여러개라면 확장자의 개수를 저장
                            getExtcount++;
                        }

                        foreach (string sh in Ex_Def)
                        { //서버에서 얻어온 시그니처를 처음부터 끝까지 비교
                            if (sh != "=")
                            {
                                string b1 = ((byte)br.ReadByte()).ToString("X2"); //바이너리 -> string
                                if (b1 == sh)
                                { //시그니처가 같다면 sh[] = {'11', '22', '33', '44'} b1 = 11 
                                    liSi2 += b1 + " ";
                                    count1++;
                                    if (count1 == count)
                                    {//시그니처 바이트 개수가 같다면 sh -> count = 4
                                        sigcount = 0;
                                        br.Close();
                                        int count2 = 0;
                                        for (int e = 0; e < getExtcount; e++)
                                        { //그 시그니처에 확장자가 여러개라면 
                                            if (lower != lotsExt[e].ToLower())
                                            { //시그니처는 같은데 확장자가 다르다면
                                                count2++;
                                                if (getExtcount == count2)
                                                { //확장자가 여러개인데 파일의 확장자와 비교하였을 때 다 같지 않다면
                                                    chColor = 1;
                                                    lipEx2 = pExtension;
                                                    lirEx2 = comExt;
                                                    liHa2 = ComputeSHA1Hash(fullFile);
                                                    ListViewItem item1 = new ListViewItem("" + liNo2);
                                                    item1.SubItems.Add(liNa2); //Extension
                                                    item1.SubItems.Add(liPa2); // FileSize
                                                    item1.SubItems.Add(lipEx2); // FileCreationDate
                                                    item1.SubItems.Add(lirEx2); // File Last Access Date
                                                    item1.SubItems.Add(liSi2); // Hd Check
                                                    item1.SubItems.Add(liHa2); // Hd Check
                                                    listView2.Items.Add(item1);
                                                    liNo2++;
                                                    break;
                                                }
                                            }
                                            else
                                                break;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        sigcount++;
                    }
                    //전체파일리스트
                    ListViewItem item2 = new ListViewItem("" + liNo1);
                    item2.SubItems.Add(liNa1);
                    item2.SubItems.Add(liSi1);
                    item2.SubItems.Add(liPa1);
                    item2.SubItems.Add(liCr1);
                    item2.SubItems.Add(liAc1);
                    item2.SubItems.Add(liHa1);
                    if (sigcount == z) //해당 파일의 확장자가 서버에 없을 때
                        item2.SubItems.Add(liReli1);
                    else
                    {
                        liReli1 = "HIGH";
                        item2.SubItems.Add(liReli1);
                    }
                    if (chColor == 1) //시그니처가 확장자와 일치하지 않을 때
                        item2.BackColor = Color.Aquamarine;
                    listView1.Items.Add(item2);
                    liNo1++;
                }
                else //파일의 크기가 20byte도 안 될때
                {
                    ListViewItem item2 = new ListViewItem("" + liNo1);
                    item2.SubItems.Add(liNa1);
                    item2.SubItems.Add(liSi1);
                    item2.SubItems.Add(liPa1);
                    item2.SubItems.Add(liCr1);
                    item2.SubItems.Add(liAc1);
                    item2.SubItems.Add(liHa1);
                    item2.SubItems.Add(liReli1);
                    item2.BackColor = Color.DarkGray;
                    listView1.Items.Add(item2);
                    liNo1++;
                }
            }
        }

        private void walkDirectory(DirectoryInfo dir)
        { //지정된 경로부터 최하위 폴더의 파일을 다 객체로 얻어옴
            string strDir = dir.ToString();
            string file = "";
            //textBox1.AppendText(strDir + " checking\n");
            //textBox1.AppendText("----------------------------------\n");
            foreach (FileInfo fInfo in dir.GetFiles())
            {
                try
                {
                    file = fInfo.ToString();
                    //textBox1.AppendText(file + '\n');
                    fCount++;
                    cpFile(fInfo);
                }
                catch (Exception ex)
                {
                    textBox1.AppendText(ex.Message + " : 오류발생" + ", file : " + file + "\n");
                    textBox1.AppendText("탐색중 입니다. 잠시 기다려주세요.\n");
                   //MessageBox.Show(ex.Message + " : 오류발생" + ", file : " + file);
                }
            }

            foreach (DirectoryInfo dInfo in dir.GetDirectories())
            {
                try
                {
                    walkDirectory(dInfo);
                }
                catch (Exception ex)
                {
                    //textBox1.AppendText("오류발생\n");
                    textBox1.AppendText(ex.Message + " : 오류발생" + ", dir : " + dInfo.Name + "\n");
                    textBox1.AppendText("탐색중 입니다. 잠시 기다려주세요.\n");
                    //MessageBox.Show(ex.Message + " : 오류발생" + ", dir : " + dInfo.Name);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        { //경로를 지정하는 폼을 띄움
            check child = new check();
            child.FormClosed += new FormClosedEventHandler(childClose);
            child.ShowDialog(this); //모달로 check폼을 열기 -> child->ShowDialog(this); 모달리스
        }

        private void childClose(object sender, FormClosedEventArgs e)
        {
            button1.Enabled = true;           
        }

        public static string ComputeSHA1Hash(string FilePath)
        {
            return ComputeHash(FilePath, new SHA1CryptoServiceProvider());
        }

        public static string ComputeHash(string FilePath, HashAlgorithm Algorithm)
        {
            FileStream FileStream = File.OpenRead(FilePath);
            try
            {
                byte[] HashResult = Algorithm.ComputeHash(FileStream);
                string ResultString = BitConverter.ToString(HashResult).Replace("-", "");
                return ResultString;
            }
            finally
            {
                FileStream.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        { //리스트에 있는 내용을 엑셀로 저장
            try
            {
                if (checkTab == 0)
                {
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                    saveFileDialog1.CreatePrompt = true;
                    saveFileDialog1.OverwritePrompt = true;

                    saveFileDialog1.DefaultExt = "*.xls";
                    saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
                    saveFileDialog1.InitialDirectory = "C:\\";

                    DialogResult result = saveFileDialog1.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        try
                        {
                            object missingType = Type.Missing;
                            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add(missingType);
                            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);
                            excelApp.Visible = false;

                            for (int i = 0; i < listView1.Items.Count; i++)
                            {
                                for (int j = 0; j < listView1.Columns.Count; j++)
                                {
                                    if (i == 0)
                                    {
                                        excelWorksheet.Cells[1, j + 1] = this.listView1.Columns[j].Text;
                                    }
                                    excelWorksheet.Cells[i + 2, j + 1] = this.listView1.Items[i].SubItems[j].Text;
                                }
                            }
                            excelBook.SaveAs(@saveFileDialog1.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, missingType, missingType, missingType, missingType, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingType, missingType, missingType, missingType, missingType);
                            excelApp.Visible = true;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                        catch
                        {
                            MessageBox.Show("Excel 파일 저장중 에러가 발생했습니다.");
                        }
                    }
                }
                else
                {
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                    saveFileDialog1.CreatePrompt = true;
                    saveFileDialog1.OverwritePrompt = true;

                    saveFileDialog1.DefaultExt = "*.xls";
                    saveFileDialog1.Filter = "Excel Files (*.xls)|*.xls";
                    saveFileDialog1.InitialDirectory = "C:\\";

                    DialogResult result = saveFileDialog1.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        try
                        {
                            object missingType = Type.Missing;
                            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add(missingType);
                            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);
                            excelApp.Visible = false;

                            for (int i = 0; i < listView2.Items.Count; i++)
                            {
                                for (int j = 0; j < listView2.Columns.Count; j++)
                                {
                                    if (i == 0)
                                    {
                                        excelWorksheet.Cells[1, j + 1] = this.listView2.Columns[j].Text;
                                    }
                                    excelWorksheet.Cells[i + 2, j + 1] = this.listView2.Items[i].SubItems[j].Text;
                                }
                            }
                            excelBook.SaveAs(@saveFileDialog1.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, missingType, missingType, missingType, missingType, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingType, missingType, missingType, missingType, missingType);
                            excelApp.Visible = true;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                        catch 
                        {
                            MessageBox.Show("Excel 파일 저장중 에러가 발생했습니다.");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        { //로그를 텍스트 문서로 저장
            SaveFileDialog saveFileDialog2 = new SaveFileDialog();

            saveFileDialog2.CreatePrompt = true;
            saveFileDialog2.OverwritePrompt = true;
            saveFileDialog2.DefaultExt = "txt";
            saveFileDialog2.Filter = "텍스트 문서 (*.txt)|*.txt";
            saveFileDialog2.InitialDirectory = System.Environment.CurrentDirectory;


            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                int x;
                string byte_buffer_string = textBox1.Text;
                string[] arrText = textBox1.Text.Split('\n');
                //int byte_buffer = Encoding.Default.GetByteCount(byte_buffer_string);
                //byte[] bData = new Byte[byte_buffer];
                FileStream sFile = new FileStream(saveFileDialog2.FileName, FileMode.Create, FileAccess.ReadWrite);
                StreamWriter SWFile = new StreamWriter(sFile);

                //텍스트박스(txtReceive)에 있는 내용을 텍스트파일로 저장하기
                for (x = 0; x < arrText.Length; x++)
                    SWFile.WriteLine(arrText[x]);
                /*bData = Encoding.Default.GetBytes(textBox_Log.Text);
                sFile.Seek(0, SeekOrigin.Begin);
                sFile.Write(bData, 0, bData.Length);*/
                SWFile.Close();
                sFile.Close();
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        { //수행 안됌...ㅜㅜㅜ
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, e);
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        { //리스트뷰에서 마우스 클릭했을 때
            if (e.Button.Equals(MouseButtons.Right))
            { //마우스 오른쪽 버튼일 때
                ListViewItem list = listView1.GetItemAt(e.X, e.Y);
                string dirpath = list.SubItems[3].Text;
                string dpath = dirpath.Substring(0, dirpath.LastIndexOf("\\"));
                string caption = list.SubItems[1].Text + " 해쉬 값";
                ContextMenu m = new ContextMenu();

                //메뉴에 들어갈 아이템을 만듭니다
                MenuItem m1 = new MenuItem();
                m1.Text = "폴더열기";
                MenuItem m2 = new MenuItem();
                m2.Text = "해쉬보기";

                m1.Click += (senders, es) =>
                {
                    System.Diagnostics.Process.Start(dpath);
                };
                m2.Click += (senders, es) =>
                {
                    liHa1 = ComputeSHA1Hash(dirpath);
                    listView1.Items[list.Index].SubItems[6].Text = liHa1;
                    //MessageBox.Show(ComputeSHA1Hash(dirpath), caption);
                };

                m.MenuItems.Add(m1);
                m.MenuItems.Add(m2);

                m.Show(listView1, new Point(e.X, e.Y));
            }
        }

        private void listView2_MouseClick(object sender, MouseEventArgs e)
        { //리스트에서 마우스를 눌렀을 때
            if (e.Button.Equals(MouseButtons.Right))
            { //마우스 오른쪽 버튼을 눌렀을 때
                ListViewItem list = listView2.GetItemAt(e.X, e.Y);
                string dirpath = list.SubItems[2].Text;
                string dpath = dirpath.Substring(0, dirpath.LastIndexOf("\\"));
                
                ContextMenu m = new ContextMenu();

                //메뉴에 들어갈 아이템을 만듭니다
                MenuItem m2 = new MenuItem();
                m2.Text = "폴더열기";

                m2.Click += (senders, es) =>
                {
                    System.Diagnostics.Process.Start(dpath);
                };

                m.MenuItems.Add(m2);

                m.Show(listView2, new Point(e.X, e.Y));
            }
        }
    }
}
