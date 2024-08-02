using FuzzySharp;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace ExcelCompare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            using ExcelPackage excel = new ExcelPackage(new FileInfo(label3.Text));
            var text = File.ReadAllText(label4.Text);
            var json = JsonConvert.DeserializeObject<List<KTB>>(text);

            var worksheet = excel.Workbook.Worksheets[0];
            progressBar1.Maximum = worksheet.Dimension.Rows - 1;
            
            for (int r = 2; r <= worksheet.Dimension.Rows; r++)
            {
                var name = $"{worksheet.Cells[r, 2].Value}";
                var star = $"{worksheet.Cells[r, 3].Value}";
                var state = $"{worksheet.Cells[r, 4].Value}";

                var result = json!.Select(x => new
                {
                    score = Fuzz.PartialRatio(TurkishCharsToEnglish(x.TesisAdi).ToUpper(), TurkishCharsToEnglish(name).ToUpper()),
                    x.BelgeTuru,
                    x.TesisAdi,
                    x.Ilce,
                    x.Sehir,
                    x.TesisSinifi,
                    x.TesisTuru
                }).OrderByDescending(x => x.score).FirstOrDefault();
            
                if (result is null) continue;
                else
                {
                    json!.RemoveAt(json.FindIndex(x => x.TesisAdi == result.TesisAdi));
                    worksheet.Cells[r, 6].Value = result.TesisAdi;
                    worksheet.Cells[r, 7].Value = $"{result.Sehir} - {result.Ilce}";
                    worksheet.Cells[r, 8].Value = $"{result.TesisTuru} - {result.TesisSinifi}";
                    worksheet.Cells[r, 9].Value = result.BelgeTuru;
                }
                progressBar1.Value = r - 1;
            }

            excel.Save();
            progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show("bitti");
        }

        public static string TurkishCharsToEnglish(string input)
        {

            return Regex.Replace(input, "[ýöüðþçÝÖÜÐÞÇ]", c =>
            {
                return c.Value switch
                {
                    "ý" => "i",
                    "ö" => "o",
                    "ü" => "u",
                    "ð" => "g",
                    "þ" => "s",
                    "ç" => "c",
                    "Ý" => "I",
                    "Ö" => "O",
                    "Ü" => "U",
                    "Ð" => "G",
                    "Þ" => "S",
                    "Ç" => "C",
                    _ => c.Value,
                };
            });
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label3.Text = "C:\\Users\\mehmetcan.kaplan\\Desktop\\sejour-hotels.xlsx";
            label4.Text = "C:\\Users\\mehmetcan.kaplan\\Desktop\\aaa.json";
        }
    }
}
