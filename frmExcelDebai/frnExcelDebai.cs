using DevExpress.DataAccess.Excel;
using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace frmExcelDebai
{
    public partial class frnExcelDebai : DevExpress.XtraEditors.XtraForm
    {
        FirestoreDb firestoreDb;
        public frnExcelDebai()
        {
            InitializeComponent();
            loaddebai();
        }

        private void frnExcelDebai_Load(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"quizzapp.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);

            firestoreDb = FirestoreDb.Create("quizz-app-7d3ec");
        }
        private List<Debai> lstdebai;
        private List<meo> lstdebais;
        private void loaddebai()
        {
            lstdebai = new List<Debai>();
            lstdebais = new List<meo>();
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select File";
            openFileDialog.Filter = "Excel (*.xlsx) | *.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Browse.Text = openFileDialog.FileName;
                ExcelDataSource excelDataSource = new ExcelDataSource();
                excelDataSource.FileName = Browse.Text;
                ExcelWorksheetSettings excelWorksheetSettings = new ExcelWorksheetSettings("Mydata", "A1:O500");
                excelDataSource.SourceOptions = new ExcelSourceOptions(excelWorksheetSettings);
                excelDataSource.SourceOptions = new CsvSourceOptions() { CellRange = "A1:O500" };
                excelDataSource.SourceOptions.SkipEmptyRows = false;
                excelDataSource.SourceOptions.UseFirstRowAsHeader = true;
                excelDataSource.Fill();
                gcData.DataSource = excelDataSource;
            }
        }
        void loaddata()
        {
            for (int i = 0; i < gvData.RowCount; i++)
            {
                Debai debai = new Debai();
                for (int j = 0; j < gvData.Columns.Count; j++)
                {
                    var value = gvData.GetRowCellValue(i, gvData.Columns[j]);
                    switch (j)
                    {
                        case 0:
                            debai.title = value.ToString();
                            break;
                        case 1:
                            debai.id = value.ToString();
                            break;
                        case 2:
                            if (value != null)
                            {
                                debai.img_url = value.ToString();
                            }
                            else { debai.img_url = ""; }
                                break;
                        case 3:
                            debai.question = value.ToString();
                            break;
                        case 4:
                            debai.indentifierA = value.ToString();
                            break;
                        case 5:
                            debai.answerA = value.ToString();
                            break;
                        case 6:
                            debai.indentifierB = value.ToString();
                            break;
                        case 7:
                            debai.answerB = value.ToString();
                            break;
                        case 8:
                            debai.indentifierC = value.ToString();
                            break;
                        case 9:
                            debai.answerC = value.ToString();
                            break;
                        case 10:
                            debai.indentifierD = value.ToString();
                            break;
                        case 11:
                            debai.answerD = value.ToString();
                            break;
                        case 12:
                            debai.conrrect_anser = value.ToString();
                            break;
                        case 13:
                            if (value != null)
                            {
                                debai.selected_answer = value.ToString();
                            }
                            else { debai.selected_answer = ""; }
                            
                            break;
                    }
                }
                lstdebai.Add(debai);
            }
            gcData.DataSource = null;
            gcData.DataSource = lstdebai;
        }
        void meoloaddata()
        {
            for (int i = 0; i < gvData.RowCount; i++)
            {
                meo meo = new meo();
                for (int j = 0; j < gvData.Columns.Count; j++)
                {
                    var value = gvData.GetRowCellValue(i, gvData.Columns[j]);
                    switch (j)
                    {
                        case 0:
                            meo.id = value.ToString();
                            break;
                        case 1:
                            meo.keys = value.ToString();
                            break;
                        case 2:
                            meo.comment = value.ToString();
                            break;
                        
                    }
                }
                lstdebais.Add(meo);
            }
            gcData.DataSource = null;
            gcData.DataSource = lstdebais;
        }

        private async void btnLuu_Click(object sender, EventArgs e)
        {
            loaddata();
            SplashScreenManager.ShowForm(typeof(WaitForm1));
            foreach (var debai in lstdebai)
            {
                string title = debai.title.Trim();
                string id = debai.id.Trim();
                string img_url = debai.img_url.Trim();
                string question = debai.question.Trim();
                string indetifierA = debai.indentifierA.Trim();
                string answerA = debai.answerA.Trim();
                string indetifierB = debai.indentifierB.Trim();
                string answerB = debai.answerB.Trim();
                string indetifierC = debai.indentifierC.Trim();
                string answerC = debai.answerC.Trim();
                string indetifierD = debai.indentifierD.Trim();
                string answerD = debai.answerD.Trim();
                string conrrect_anser = debai.conrrect_anser.Trim();
                string selected_answer = debai.selected_answer.Trim();
                WriteBatch batch = firestoreDb.StartBatch();
                //DocumentReference docRef1 = firestoreDb.Collection("setOfTopics").Document(txta.Text);
                //Dictionary<string, object> keys1 = new Dictionary<string, object>()
                // {
                //    {"title",txtTitle.Text},
                // };
                //await docRef1.SetAsync(keys1);
                //DocumentReference docRef = firestoreDb.Collection("setOfTopics").Document(txta.Text).Collection("questions").Document(id.Trim());
                //Dictionary<string, object> keys = new Dictionary<string, object>()
                // {
                //    {"selected_answer",selected_answer},
                //     {"correct_answer",conrrect_anser},
                //     {"image_url",img_url },
                //     {"question",question},
                //     {"id_title",title }
                // };
                DocumentReference docRef = firestoreDb.Collection(txta.Text).Document(id.Trim());
                Dictionary<string, object> keys = new Dictionary<string, object>()
                 {
                    {"selected_answer",selected_answer},
                     {"correct_answer",conrrect_anser},
                     {"image_url",img_url },
                     {"question",question},
                     {"id_title",title }
                 };
                await docRef.SetAsync(keys);
                DocumentReference subCollection1 = docRef.Collection("answers").Document(indetifierA.Trim());
                Dictionary<string, object> keyss1 = new Dictionary<string, object>()
                     {
                         {"identifier",indetifierA },
                         {"answer",answerA },
                     };

                    await subCollection1.SetAsync(keyss1);

                DocumentReference subCollection2 = docRef.Collection("answers").Document(indetifierB.Trim());
                Dictionary<string, object> keyss2 = new Dictionary<string, object>()
                     {
                         {"identifier",indetifierB },
                         {"answer",answerB },
                     };

                    await subCollection2.SetAsync(keyss2);

                DocumentReference subCollection3 = docRef.Collection("answers").Document(indetifierC.Trim());
                Dictionary<string, object> keyss3 = new Dictionary<string, object>()
                     {
                         {"identifier",indetifierC },
                         {"answer",answerC },
                     };

                    await subCollection3.SetAsync(keyss3);

                DocumentReference subCollection4 = docRef.Collection("answers").Document(indetifierD.Trim());
                Dictionary<string, object> keyss4 = new Dictionary<string, object>()
                     {
                         {"identifier",indetifierD },
                         {"answer",answerD },
                     };

                    await subCollection4.SetAsync(keyss4);
  
            }
            SplashScreenManager.CloseForm();
            MessageBox.Show("Thêm Dữ Liệu Thành Công?", "Thông báo", MessageBoxButtons.OKCancel);
        }

       
            

        private async void simpleButton1_Click(object sender, EventArgs e)
        {
                meoloaddata();
                SplashScreenManager.ShowForm(typeof(WaitForm1));
                foreach (var meo in lstdebais)
                {
                    string id = meo.id.Trim();
                    string keysmeo = meo.keys.Trim();
                    string comment = meo.comment.Trim();
                    WriteBatch batch = firestoreDb.StartBatch();
                    DocumentReference docRef1 = firestoreDb.Collection("tips").Document(txta.Text);
                    Dictionary<string, object> keys1 = new Dictionary<string, object>()
                 {
                    {"title",txta.Text},
                 };
                    await docRef1.SetAsync(keys1);
                    DocumentReference docRef = firestoreDb.Collection("tips").Document(txta.Text).Collection("keyboard").Document(id.Trim());
                    Dictionary<string, object> keys = new Dictionary<string, object>()
                 {
                    {"title",txtTitle.Text},
                 };
                    await docRef.SetAsync(keys);
                    DocumentReference docRef2 = firestoreDb.Collection("tips").Document(txta.Text).Collection("keyboard").Document(id.Trim()).Collection("shortcuts").Document();
                    Dictionary<string, object> keyss = new Dictionary<string, object>()
                 {
                    {"comment",comment},
                    {"keys",keysmeo},
                 };
                    await docRef2.SetAsync(keyss);

                }
                SplashScreenManager.CloseForm();
                MessageBox.Show("Thêm Dữ Liệu Thành Công?", "Thông báo", MessageBoxButtons.OKCancel);
            }
        }
    
}