using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BSIImportExcel
{
    public partial class Form1 : Form
    {
        int _fstRow = 2;
        int _colAgencia = 1;
        int _colConta = 2;
        string _sqlString;
        string _connString;

        public Form1()
        {
            InitializeComponent();

            txtOrigem.Text = ConfigurationManager.AppSettings["defaultDir"];
            _connString = ConfigurationManager.ConnectionStrings["accessConnection"].ConnectionString;
            _sqlString = getCmdText();
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            importarArquivo();
        }

        private void importarArquivo()
        {
            bool processado;

            foreach (var item in getArquivos())
            {
                Excel.Application appExcel = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet workSheet = null;
                Excel.Range range;
                processado = false;

                try
                {
                    appExcel = new Excel.Application();
                    workbook = appExcel.Workbooks.Open(item, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    workSheet = (Excel.Worksheet)workbook.Worksheets[1];
                    range = workSheet.UsedRange;

                    for (int i = _fstRow; i <= range.Rows.Count; i++)
                    {
                        if ((!string.IsNullOrEmpty((range.Cells[i, _colAgencia] as Excel.Range).Value.ToString())) &&
                            !string.IsNullOrEmpty((range.Cells[i,_colConta] as Excel.Range).Value.ToString()))
                        {
                            // TODO: insert into access table
                            Registro registro = new Registro();
                            registro.Agencia = (range.Cells[i, _colAgencia] as Excel.Range).Value.ToString();
                            registro.Conta = ((Excel.Range)range.Cells[i, _colConta]).Value.ToString();

                            // voltar aki
                            // updateDataBase(registro);
                        }
                    }

                    processado = true;
                }
                catch (Exception e)
                {
                    // TODO: log
                    MessageBox.Show(e.Message);
                }
                finally
                {
                    workbook.Close(true, null, null);
                    appExcel.Quit();

                    releaseObject(workSheet);
                    releaseObject(workbook);
                    releaseObject(appExcel);

                    if (processado == true)
                    {
                        moverArquivoProcessado(item);
                    }
                }
            }
        }

        private void moverArquivoProcessado(string arquivo)
        {
            FileInfo file = new FileInfo(arquivo);
            string processado = Path.Combine(txtOrigem.Text, "Processados", file.Name);

            if (File.Exists(processado))
            {
                File.Delete(processado);
            }

            try
            {
                file.MoveTo(processado);
            }
            catch (Exception)
            {
                // TODO: log falha em mover arquivo processado
                MessageBox.Show("Falha em mover arquivo processado");
            }
        }

        private List<string> getArquivos()
        {
            if (string.IsNullOrEmpty(txtOrigem.Text))
            {
                throw new ArgumentException("Diretório origem inválido");
            }

            if (!Directory.Exists(txtOrigem.Text))
            {
                throw new DirectoryNotFoundException("Diretório origem inexistente");
            }

            return Directory.GetFiles(txtOrigem.Text.Trim())
                .Where(a => !a.Contains("~$") && (a.ToLower().EndsWith(".xls") || a.ToLower().EndsWith(".xlsx"))).ToList();
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                // TODO: log release object error
            }
            finally
            {
                GC.Collect();
            }
        }

        private void updateDataBase(Registro reg)
        {
            using (OleDbConnection conn = new OleDbConnection(_connString))
            {
                conn.Open();
                using (OleDbCommand cmd = new OleDbCommand(_sqlString, conn))
                {
                    cmd.Parameters.AddWithValue("@Agencia", reg.Agencia);
                    cmd.Parameters.AddWithValue("@Conta", reg.Conta);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private string getCmdText()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("INSERT INTO Tabela (Agencia, Conta)")
                .Append(" VALUES (@Agencia,@Conta)");

            return sb.ToString();
        }

        private void tmImport_Tick(object sender, EventArgs e)
        {
            importarArquivo();
        }
    }

    public class Registro
    {
        public string Agencia { get; set; }
        public string Conta { get; set; }
    }
}
