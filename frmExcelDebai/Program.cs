using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace frmExcelDebai
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            FirestoreDb firestoreDb;
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + @"quizzapp.json";
                Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);

                firestoreDb = FirestoreDb.Create("quizz-app-7d3ec");
                MessageBox.Show("Kết Nối Thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kết Nối Thất Bại" + ex);
            }
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frnExcelDebai());
        }
    }
}
