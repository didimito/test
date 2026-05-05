using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;

namespace Sinematek_Foto_Poster
{
    public partial class Form1 : Form
    {
        private List<(string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan)> allFilms = new();
        private List<(string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan)> currentFilms = new();
        private (string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan)? currentFilm = null;

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            try
            {
                LoadAllFilms();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading films: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAllFilms()
        {
            allFilms = DatabaseHelper.GetAllFilms();
            currentFilms = new List<(string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan)>(allFilms);
            PopulateDataGridView(currentFilms);

            // Auto-select the first item to show poster
            if (dataGridViewSearchResults.Rows.Count > 0)
            {
                dataGridViewSearchResults.Rows[0].Selected = true;
            }
        }

        private void PopulateDataGridView(List<(string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan)> films)
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("No Penempatan", typeof(string));
            dataTable.Columns.Add("Judul Film", typeof(string));
            dataTable.Columns.Add("Produksi", typeof(string));
            dataTable.Columns.Add("Tahun Produksi", typeof(int));

            foreach (var film in films)
            {
                dataTable.Rows.Add(film.noPenempatan, film.judulFilm, film.produksi, film.tahunProduksi);
            }

            dataGridViewSearchResults.DataSource = dataTable;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string searchTerm = txtSearchLeft.Text.Trim();
                if (string.IsNullOrEmpty(searchTerm))
                {
                    LoadAllFilms();
                    return;
                }

                currentFilms = DatabaseHelper.SearchFilmsLike(searchTerm);
                PopulateDataGridView(currentFilms);

                // Auto-select first result
                if (dataGridViewSearchResults.Rows.Count > 0)
                {
                    dataGridViewSearchResults.Rows[0].Selected = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error searching: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridViewSearchResults_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewSearchResults.SelectedRows.Count == 0)
                {
                    ClearDisplayFields();
                    currentFilm = null;
                    SetDetailFieldsEditable(false);
                    return;
                }

                var selectedRow = dataGridViewSearchResults.SelectedRows[0];
                int selectedIndex = selectedRow.Index;
                var selectedFilm = currentFilms[selectedIndex];
                currentFilm = selectedFilm;

                // Get full details from database using the hidden record ID and title
                var fullDetails = DatabaseHelper.SearchFilmsExact(selectedFilm.judulFilm, selectedFilm.recordId);
                if (fullDetails.HasValue)
                {
                    DisplayFilmDetails(fullDetails.Value);
                    LoadFilmPoster(selectedFilm.poster);
                }
                SetDetailFieldsEditable(false); // Make fields read-only when displaying
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error displaying details: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayFilmDetails((string recordId, string noPenempatan, string judulFilm, string produksi, int tahunProduksi, string poster, string negara, string bahasa, string warna, string jenisFilm, string ukuranReel, string jumlahReel, string panjangReel, string masaPutar, string rak, string copy, string keterangan) film)
        {
            txtNoPenempatan.Text = film.noPenempatan;
            txtJudulFilm.Text = film.judulFilm;
            txtProduksi.Text = film.produksi;
            txtTahunProduksi.Text = film.tahunProduksi.ToString();
            txtNegara.Text = film.negara ?? "";
            txtBahasa.Text = film.bahasa ?? "";
            txtWarna.Text = film.warna ?? "";
            txtJenisFilm.Text = film.jenisFilm ?? "";
            txtUkuranReel.Text = film.ukuranReel ?? "";
            txtJumlahReel.Text = film.jumlahReel ?? "";
            txtPanjangReel.Text = film.panjangReel ?? "";
            txtMasaPutar.Text = film.masaPutar ?? "";
            txtRak.Text = film.rak ?? "";
            txtCopy.Text = film.copy ?? "";
            txtKeterangan.Text = film.keterangan ?? "";
        }

        private void ClearDisplayFields()
        {
            txtNoPenempatan.Text = "";
            txtJudulFilm.Text = "";
            txtProduksi.Text = "";
            txtTahunProduksi.Text = "";
            txtNegara.Text = "";
            txtBahasa.Text = "";
            txtWarna.Text = "";
            txtJenisFilm.Text = "";
            txtUkuranReel.Text = "";
            txtJumlahReel.Text = "";
            txtPanjangReel.Text = "";
            txtMasaPutar.Text = "";
            txtRak.Text = "";
            txtCopy.Text = "";
            txtKeterangan.Text = "";
            pictureBox1.Image = null;
        }

        private void LoadFilmPoster(string posterPath)
        {
            try
            {
                Debug.WriteLine($"LoadFilmPoster called. posterPath='{posterPath}'");
                if (string.IsNullOrWhiteSpace(posterPath))
                {
                    pictureBox1.Image?.Dispose();
                    pictureBox1.Image = null;
                    return;
                }

                posterPath = posterPath.Trim();
                if (posterPath.StartsWith("\"", StringComparison.Ordinal) && posterPath.EndsWith("\"", StringComparison.Ordinal))
                {
                    posterPath = posterPath[1..^1].Trim();
                }

                var separators = new[] { '#', ';', '|' };
                var parts = posterPath.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => p.Trim())
                    .Where(p => !string.IsNullOrEmpty(p))
                    .ToArray();
                if (parts.Length > 0)
                {
                    posterPath = parts[0];
                    foreach (var part in parts)
                    {
                        var candidate = Path.IsPathRooted(part) ? part : Path.Combine(Application.StartupPath, part);
                        if (File.Exists(candidate))
                        {
                            posterPath = part;
                            break;
                        }
                    }
                }

                if (posterPath.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    posterPath = new Uri(posterPath).LocalPath;
                    Debug.WriteLine($"Converted to local path: '{posterPath}'");
                }

                if (posterPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    posterPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    pictureBox1.Image?.Dispose();
                    pictureBox1.Image = null;
                    pictureBox1.LoadAsync(posterPath);
                    pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                    return;
                }

                string abs = Path.IsPathRooted(posterPath)
                    ? posterPath
                    : Path.Combine(Application.StartupPath, posterPath);

                if (!File.Exists(abs))
                {
                    pictureBox1.Image?.Dispose();
                    pictureBox1.Image = null;
                    Debug.WriteLine($"Poster file not found: {abs}");
                    MessageBox.Show($"Poster file not found:\n{abs}", "Poster Load Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                pictureBox1.Image?.Dispose();
                pictureBox1.Image = new Bitmap(abs);
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox1.Refresh();
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
                Debug.WriteLine($"Error loading poster: {ex}");
                MessageBox.Show($"Error loading poster:\n{ex.Message}", "Poster Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // Clear all fields and make them editable for new film entry
            ClearDisplayFields();
            SetDetailFieldsEditable(true);
            currentFilm = null; // No current film selected
            MessageBox.Show("Enter film details and click Update to save.", "Add New Film", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (currentFilm == null && !AreDetailFieldsEditable())
            {
                MessageBox.Show("Please select a film from the list or click Add to create a new film.", "No Film Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string noPenempatan = txtNoPenempatan.Text.Trim();
                string judulFilm = txtJudulFilm.Text.Trim();
                string produksi = txtProduksi.Text.Trim();
                string tahunProduksi = txtTahunProduksi.Text.Trim();
                string negara = txtNegara.Text.Trim();
                string bahasa = txtBahasa.Text.Trim();
                string warna = txtWarna.Text.Trim();
                string jenisFilm = txtJenisFilm.Text.Trim();
                string ukuranReel = txtUkuranReel.Text.Trim();
                string jumlahReel = txtJumlahReel.Text.Trim();
                string panjangReel = txtPanjangReel.Text.Trim();
                string masaPutar = txtMasaPutar.Text.Trim();
                string rak = txtRak.Text.Trim();
                string copy = txtCopy.Text.Trim();
                string keterangan = txtKeterangan.Text.Trim();

                if (string.IsNullOrEmpty(judulFilm))
                {
                    MessageBox.Show("Judul Film is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (currentFilm == null)
                {
                    // Adding new film
                    DatabaseHelper.AddFilm(noPenempatan, judulFilm, produksi, tahunProduksi, "", negara, bahasa, warna, jenisFilm, ukuranReel, jumlahReel, panjangReel, masaPutar, rak, copy, keterangan);
                    MessageBox.Show("Film added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Updating existing film
                    DatabaseHelper.UpdateFilm(currentFilm.Value.recordId, noPenempatan, judulFilm, produksi, tahunProduksi, currentFilm.Value.poster, negara, bahasa, warna, jenisFilm, ukuranReel, jumlahReel, panjangReel, masaPutar, rak, copy, keterangan);
                    MessageBox.Show("Film updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // Refresh the list and clear selection
                LoadAllFilms();
                dataGridViewSearchResults.ClearSelection();
                ClearDisplayFields();
                SetDetailFieldsEditable(false);
                currentFilm = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving film: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (currentFilm == null)
            {
                MessageBox.Show("Please select a film from the list to delete.", "No Film Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show($"Are you sure you want to delete '{currentFilm.Value.judulFilm}'?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    DatabaseHelper.DeleteFilm(currentFilm.Value.recordId);
                    MessageBox.Show("Film deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadAllFilms();
                    dataGridViewSearchResults.ClearSelection();
                    ClearDisplayFields();
                    currentFilm = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting film: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SetDetailFieldsEditable(bool editable)
        {
            txtNoPenempatan.ReadOnly = !editable;
            txtJudulFilm.ReadOnly = !editable;
            txtProduksi.ReadOnly = !editable;
            txtTahunProduksi.ReadOnly = !editable;
            txtNegara.ReadOnly = !editable;
            txtBahasa.ReadOnly = !editable;
            txtWarna.ReadOnly = !editable;
            txtJenisFilm.ReadOnly = !editable;
            txtUkuranReel.ReadOnly = !editable;
            txtJumlahReel.ReadOnly = !editable;
            txtPanjangReel.ReadOnly = !editable;
            txtMasaPutar.ReadOnly = !editable;
            txtRak.ReadOnly = !editable;
            txtCopy.ReadOnly = !editable;
            txtKeterangan.ReadOnly = !editable;
        }

        private bool AreDetailFieldsEditable()
        {
            return !txtNoPenempatan.ReadOnly;
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            Bitmap captured = null;

            try
            {
                string outputPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "FormExport.pdf");

                Color bg = this.BackColor;
                int tolerance = 24;

                // Capture on UI thread (DrawToBitmap must be called on UI thread)
                var clientSize = this.ClientSize;
                captured = new Bitmap(clientSize.Width, clientSize.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                this.DrawToBitmap(captured, new Rectangle(0, 0, clientSize.Width, clientSize.Height));

                // Process heavy work on background thread (must not touch UI controls)
                await System.Threading.Tasks.Task.Run(() =>
                {
                   
                });

                MessageBox.Show($"Exported PDF to: {outputPath}", "Export complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button3.Enabled = true;
                // Dispose captured bitmap
                captured?.Dispose();
            }
        }
    }
}
