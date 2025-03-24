using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Adobe.PDFServicesSDK;
using Adobe.PDFServicesSDK.auth;
using Adobe.PDFServicesSDK.io;
using Adobe.PDFServicesSDK.pdfjobs.jobs;
using Adobe.PDFServicesSDK.pdfjobs.parameters.compresspdf;
using Adobe.PDFServicesSDK.pdfjobs.results;
using log4net;
using log4net.Config;
using log4net.Repository;
using System.Net;
using CompressionLevel = Adobe.PDFServicesSDK.pdfjobs.parameters.compresspdf.CompressionLevel;
using ICredentials = Adobe.PDFServicesSDK.auth.ICredentials;
using Newtonsoft.Json;

namespace CompressPDFAdobe
{
    class Config
    {
        public PDFServicesConfig PDFServices { get; set; }
    }

    class PDFServicesConfig
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
    }

    class ConfigurationManager
    {
        public static Config LoadConfig(string filePath = "config.json")
        {
            // Obtener la ruta base del directorio de la aplicación
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Construir la ruta completa al archivo de configuración
            string fullPath = Path.Combine(baseDirectory, filePath);

            // Verificar si el archivo existe
            if (File.Exists(fullPath))
            {
                string jsonContent = File.ReadAllText(fullPath);
                return JsonConvert.DeserializeObject<Config>(jsonContent);
            }
            throw new FileNotFoundException($"El archivo de configuración no se encuentra en la ruta: {fullPath}");
        }

    }

    public partial class MainForm : Form
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(MainForm));
        private ListBox lstFiles;
        private Label lblFolderPath;
        private TrackBar trkCompressionLevel;
        private Button btnSelectFolder;
        private Button btnCompress;
        private string selectedFolderPath;


        private DataGridView dgvFiles;
        private ProgressBar progressBar;

        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            dgvFiles = new DataGridView();
            FileName = new DataGridViewTextBoxColumn();
            Status = new DataGridViewTextBoxColumn();
            Mensaje = new DataGridViewTextBoxColumn();
            FileNameResult = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn1 = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn2 = new DataGridViewTextBoxColumn();
            lblFolderPath = new Label();
            trkCompressionLevel = new TrackBar();
            btnSelectFolder = new Button();
            btnCompress = new Button();
            progressBar = new ProgressBar();
            label1 = new Label();
            label2 = new Label();
            panel1 = new Panel();
            lBProcesados = new Label();
            Label5 = new Label();
            lBCantidad = new Label();
            label3 = new Label();
            ((System.ComponentModel.ISupportInitialize)dgvFiles).BeginInit();
            ((System.ComponentModel.ISupportInitialize)trkCompressionLevel).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // dgvFiles
            // 
            dgvFiles.AllowUserToAddRows = false;
            dgvFiles.AllowUserToDeleteRows = false;
            dgvFiles.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvFiles.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvFiles.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvFiles.Columns.AddRange(new DataGridViewColumn[] { FileName, Status, Mensaje, FileNameResult });
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.Window;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F);
            dataGridViewCellStyle1.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dgvFiles.DefaultCellStyle = dataGridViewCellStyle1;
            dgvFiles.Location = new Point(12, 88);
            dgvFiles.Name = "dgvFiles";
            dgvFiles.ReadOnly = true;
            dgvFiles.RowHeadersVisible = false;
            dgvFiles.RowHeadersWidth = 62;
            dgvFiles.Size = new Size(1082, 203);
            dgvFiles.TabIndex = 1;
            // 
            // FileName
            // 
            FileName.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            FileName.FillWeight = 30F;
            FileName.HeaderText = "Nombre de fichero";
            FileName.MinimumWidth = 8;
            FileName.Name = "FileName";
            FileName.ReadOnly = true;
            FileName.ToolTipText = "Nombre del Fichero";
            // 
            // Status
            // 
            Status.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Status.FillWeight = 10F;
            Status.HeaderText = "Estado";
            Status.MinimumWidth = 8;
            Status.Name = "Status";
            Status.ReadOnly = true;
            Status.ToolTipText = "Estado";
            // 
            // Mensaje
            // 
            Mensaje.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Mensaje.FillWeight = 30F;
            Mensaje.HeaderText = "Descripción";
            Mensaje.MinimumWidth = 8;
            Mensaje.Name = "Mensaje";
            Mensaje.ReadOnly = true;
            Mensaje.ToolTipText = "Descripción";
            // 
            // FileNameResult
            // 
            FileNameResult.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            FileNameResult.FillWeight = 30F;
            FileNameResult.HeaderText = "Fichero resultado";
            FileNameResult.MinimumWidth = 8;
            FileNameResult.Name = "FileNameResult";
            FileNameResult.ReadOnly = true;
            FileNameResult.ToolTipText = "Fichero resultado";
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewTextBoxColumn1.HeaderText = "Archivo PDF";
            dataGridViewTextBoxColumn1.MinimumWidth = 8;
            dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            dataGridViewTextBoxColumn1.Width = 150;
            // 
            // dataGridViewTextBoxColumn2
            // 
            dataGridViewTextBoxColumn2.HeaderText = "Estado";
            dataGridViewTextBoxColumn2.MinimumWidth = 8;
            dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            dataGridViewTextBoxColumn2.Width = 150;
            // 
            // lblFolderPath
            // 
            lblFolderPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblFolderPath.Location = new Point(12, 9);
            lblFolderPath.Name = "lblFolderPath";
            lblFolderPath.Size = new Size(1082, 76);
            lblFolderPath.TabIndex = 0;
            lblFolderPath.Text = "Carpeta seleccionada: Ninguna";
            // 
            // trkCompressionLevel
            // 
            trkCompressionLevel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            trkCompressionLevel.Location = new Point(837, 297);
            trkCompressionLevel.Minimum = 1;
            trkCompressionLevel.Name = "trkCompressionLevel";
            trkCompressionLevel.Size = new Size(257, 69);
            trkCompressionLevel.TabIndex = 2;
            trkCompressionLevel.Value = 5;
            // 
            // btnSelectFolder
            // 
            btnSelectFolder.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSelectFolder.Location = new Point(12, 414);
            btnSelectFolder.Name = "btnSelectFolder";
            btnSelectFolder.Size = new Size(382, 73);
            btnSelectFolder.TabIndex = 3;
            btnSelectFolder.Text = "Seleccionar Carpeta";
            btnSelectFolder.Click += btnSelectFolder_Click;
            // 
            // btnCompress
            // 
            btnCompress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCompress.Location = new Point(740, 414);
            btnCompress.Name = "btnCompress";
            btnCompress.Size = new Size(354, 73);
            btnCompress.TabIndex = 4;
            btnCompress.Text = "Comprimir PDFs";
            btnCompress.Click += btnCompress_Click;
            // 
            // progressBar
            // 
            progressBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new Point(12, 388);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(1082, 20);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.TabIndex = 0;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            label1.AutoSize = true;
            label1.Location = new Point(852, 341);
            label1.Name = "label1";
            label1.Size = new Size(46, 25);
            label1.TabIndex = 5;
            label1.Text = "Bajo";
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            label2.AutoSize = true;
            label2.Location = new Point(1049, 341);
            label2.Name = "label2";
            label2.Size = new Size(45, 25);
            label2.TabIndex = 6;
            label2.Text = "Alto";
            // 
            // panel1
            // 
            panel1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            panel1.Controls.Add(lBProcesados);
            panel1.Controls.Add(Label5);
            panel1.Controls.Add(lBCantidad);
            panel1.Controls.Add(label3);
            panel1.Location = new Point(12, 297);
            panel1.Name = "panel1";
            panel1.Size = new Size(819, 85);
            panel1.TabIndex = 7;
            // 
            // lBProcesados
            // 
            lBProcesados.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            lBProcesados.AutoSize = true;
            lBProcesados.Location = new Point(123, 44);
            lBProcesados.Name = "lBProcesados";
            lBProcesados.Size = new Size(22, 25);
            lBProcesados.TabIndex = 4;
            lBProcesados.Text = "0";
            // 
            // Label5
            // 
            Label5.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Label5.AutoSize = true;
            Label5.Location = new Point(14, 44);
            Label5.Name = "Label5";
            Label5.Size = new Size(107, 25);
            Label5.TabIndex = 2;
            Label5.Text = "Procesados:";
            // 
            // lBCantidad
            // 
            lBCantidad.AutoSize = true;
            lBCantidad.Location = new Point(107, 9);
            lBCantidad.Name = "lBCantidad";
            lBCantidad.Size = new Size(22, 25);
            lBCantidad.TabIndex = 1;
            lBCantidad.Text = "0";
            lBCantidad.TextAlign = ContentAlignment.TopCenter;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(14, 9);
            label3.Name = "label3";
            label3.Size = new Size(87, 25);
            label3.TabIndex = 0;
            label3.Text = "Cantidad:";
            // 
            // MainForm
            // 
            ClientSize = new Size(1106, 499);
            Controls.Add(panel1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(progressBar);
            Controls.Add(lblFolderPath);
            Controls.Add(dgvFiles);
            Controls.Add(trkCompressionLevel);
            Controls.Add(btnSelectFolder);
            Controls.Add(btnCompress);
            Name = "MainForm";
            Text = "PDF Compressor";
            ((System.ComponentModel.ISupportInitialize)dgvFiles).EndInit();
            ((System.ComponentModel.ISupportInitialize)trkCompressionLevel).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = dialog.SelectedPath;
                    lblFolderPath.Text = $"Carpeta seleccionada: {selectedFolderPath}";

                    try
                    {
                        string[] pdfFiles = Directory.GetFiles(selectedFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                        dgvFiles.Rows.Clear(); // Limpiamos antes de cargar nuevos archivos

                        if (pdfFiles.Length == 0)
                        {
                            MessageBox.Show("No se encontraron archivos PDF en la carpeta seleccionada.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        foreach (var file in pdfFiles)
                        {
                            dgvFiles.Rows.Add(Path.GetFileName(file), "Pendiente");
                        }
                        lBCantidad.Text = pdfFiles.Length.ToString();
                        lBProcesados.Text = $"0 de {pdfFiles.Length}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al leer los archivos PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }



        private void btnCompress_Click(object sender, EventArgs e)
        {
            _ = CompressSelectedFiles();
        }

        private async Task CompressSelectedFiles()
        {
            btnCompress.Enabled = true;
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Maximum = dgvFiles.Rows.Count;

            await Task.Run(() =>
            {
                try
                {
                    Config config = ConfigurationManager.LoadConfig();
                    string clientId = config.PDFServices.ClientId;
                    string clientSecret = config.PDFServices.ClientSecret;

                    if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                    {
                        MessageBox.Show("Las credenciales de Adobe no están configuradas.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    ICredentials credentials = new ServicePrincipalCredentials(clientId, clientSecret);
                    PDFServices pdfServices = new PDFServices(credentials);

                    for (int i = 0; i < dgvFiles.Rows.Count; i++)
                    {
                        string fileName = string.Empty;


                        dgvFiles.Invoke(() =>
                        {
                            dgvFiles.Rows[i].Cells["Status"].Value = "En Proceso";
                            dgvFiles.ClearSelection();
                            dgvFiles.Rows[i].Selected = true;
                            
                            // Solo hacemos scroll si la fila no está a la vista
                            if (i >= dgvFiles.FirstDisplayedScrollingRowIndex + dgvFiles.DisplayedRowCount(false) ||
                                i < dgvFiles.FirstDisplayedScrollingRowIndex)
                            {
                                dgvFiles.FirstDisplayedScrollingRowIndex = i;
                            }
                        });

                        Application.DoEvents();

                        try
                        {
                            fileName = dgvFiles.Invoke(() => dgvFiles.Rows[i].Cells["FileName"].Value.ToString());

                            Application.DoEvents();

                            string filePath = Path.Combine(selectedFolderPath, fileName);

                            long originalSize = new FileInfo(filePath).Length;

                            using Stream inputStream = File.OpenRead(filePath);
                            IAsset asset = pdfServices.Upload(inputStream, PDFServicesMediaType.PDF.GetMIMETypeValue());

                            CompressionLevel compressionLevel = MapCompressionLevel(trkCompressionLevel.Invoke(() => trkCompressionLevel.Value));
                            CompressPDFParams compressParams = new CompressPDFParams.Builder()
                                .WithCompressionLevel(compressionLevel)
                                .Build();

                            CompressPDFJob compressJob = new CompressPDFJob(asset);
                            compressJob.SetParams(compressParams);

                            string location = pdfServices.Submit(compressJob);
                            var response = pdfServices.GetJobResult<CompressPDFResult>(location, typeof(CompressPDFResult));

                            IAsset resultAsset = response.Result.Asset;
                            StreamAsset streamAsset = pdfServices.GetContent(resultAsset);

                            string outputFilePath = Path.Combine(selectedFolderPath, Path.GetFileNameWithoutExtension(filePath) + "_compressed.pdf");

                            using FileStream outputStream = File.Create(outputFilePath);
                            streamAsset.Stream.CopyTo(outputStream);

                            long compressedSize = new FileInfo(outputFilePath).Length;
                            double reduction = 100 - ((compressedSize * 100.0) / originalSize);

                            dgvFiles.Invoke(() =>
                            {
                                dgvFiles.Rows[i].Cells["Status"].Value = "Comprimido";
                                dgvFiles.Rows[i].Cells["Mensaje"].Value = $"Fichero original: {FormatSize(originalSize)}. Fichero reducido: {FormatSize(compressedSize)}. % Reducción: {reduction} ";
                                dgvFiles.Rows[i].Cells["FileNameResult"].Value = $"{outputFilePath}";
                               
                                //progressBar.PerformStep();
                            });
                            Application.DoEvents();
                        }
                        catch (Exception exFile)
                        {
                            log.Error($"Error al comprimir el archivo {fileName}", exFile);
                            dgvFiles.Invoke(() =>
                            {
                                dgvFiles.Rows[i].Cells["Status"].Value = "Error";
                                dgvFiles.Rows[i].Cells["Mensaje"].Value = exFile.Message;
                                //progressBar.PerformStep();
                            });
                            Application.DoEvents();
                        }

                        this.Invoke(() =>
                        {
                            progressBar.Value = i + 1;
                            var lblProgreso = Controls.Find("lBProcesados", true).FirstOrDefault() as Label;
                            if (lblProgreso != null)
                            {
                                lblProgreso.Text = $"{i + 1} de {dgvFiles.Rows.Count}";
                            }
                        });
                        Application.DoEvents();

                    }

                    MessageBox.Show("Proceso completado.", "Resultado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    log.Error("Error general en la compresión de archivos PDF", ex);
                    MessageBox.Show("Error en la compresión de archivos PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });

            btnCompress.Enabled = true;
        }

        string FormatSize(long bytes)
        {
            if (bytes >= 1024 * 1024)
                return $"{bytes / (1024.0 * 1024):0.##} MB";
            else
                return $"{bytes / 1024.0:0.##} KB";
        }


        private static CompressionLevel MapCompressionLevel(int value)
        {
            return value switch
            {
                <= 3 => CompressionLevel.LOW,
                <= 7 => CompressionLevel.MEDIUM,
                _ => CompressionLevel.HIGH,
            };
        }

        private static void ConfigureLogging()
        {
            ILoggerRepository logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));
        }

        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private DataGridViewTextBoxColumn FileName;
        private DataGridViewTextBoxColumn Status;
        private DataGridViewTextBoxColumn Mensaje;
        private DataGridViewTextBoxColumn FileNameResult;
        private Label label1;
        private Label label2;
        private Panel panel1;
        private Label lBProcesados;
        private Label Label5;
        private Label lBCantidad;
        private Label label3;
    }
}
