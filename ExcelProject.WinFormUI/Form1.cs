using ExcelProject.Business;
using ExcelProject.DataAccess.Concrete.Repository;
using ExcelProject.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.WinFormUI
{
    public partial class Form1 : Form
    {
        FileManager _fileService = new FileManager(new FileRepository());
        private List<T_GK_MAIN> _tgkMainList = null;
        private T_GK_MAIN _currentPage = null;
        public Form1()
        {
            try
            {
                InitializeComponent();
                LoadUI();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Form yüklenirken hata oluştu. Hata içeriği: " + ex.Message);
            }
        }

        public void LoadUI()
        {
            //_tgkMainList = _fileService.Read(Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory())) + "\\ornek_veri.xlsx");
            _tgkMainList = _fileService.Read(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()) + "\\ornek_veri.xlsx");
            if (_tgkMainList == null)
                return;
            SetCurrentPage(1);
            GetSelectedPageItems();
        }

        public void SetCurrentPage(int pageId)
        {
            if (_tgkMainList == null)
                return;
            _currentPage = _tgkMainList.FirstOrDefault(p => p.id == pageId);
        }

        public void GetSelectedPageItems()
        {
            listView1.Items.Clear();
            foreach (var item in _currentPage.allItems)
                listView1.Items.Add(new ListViewItem(new string[] { (item.id).ToString(), item.code1, item.code2, item.code3, item.code4, item.code5 }));
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            SetCurrentPage(1);
            GetSelectedPageItems();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            SetCurrentPage(2);
            GetSelectedPageItems();
        }
    }
}
