using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing.Imaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace PdfConvertImage
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Load();
        }

        private void Load()
        {

            this.pdfSouceLb.Content = Directory.GetCurrentDirectory() + @"\Samples";
            this.wordOutPutLb.Content = Directory.GetCurrentDirectory() + @"\Samples";
        }

        private void ConvertImage_Click(object sender, RoutedEventArgs e)
        {
            //获取文件目录
            var projectDir = this.pdfSouceLb.Content.ToString();
            //获取所有的pdf文件
            var array = Directory.GetFiles(projectDir, "*.pdf");
            if (array.Length <= 0) {
                MessageBox.Show(projectDir+"  不存在pdf文件");
                return;
            }

            foreach (var file in array)
            {
                //获取文件名称
                string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                if (!Directory.Exists(projectDir + @"\" + fileName)) Directory.CreateDirectory(projectDir + @"\" + fileName);
                try
                {
                    using (var document = PdfDocument.Load(file))
                    {
                        var pageCount = document.PageCount;

                        for (int i = 0; i < pageCount; i++)
                        {
                            var dpi = 100;

                            using (var image = document.Render(i, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                            {
                                var encoder = ImageCodecInfo.GetImageEncoders()
                                    .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                                var encParams = new EncoderParameters(1);
                                encParams.Param[0] = new EncoderParameter(
                                    System.Drawing.Imaging.Encoder.Quality, 100L);

                                image.Save(projectDir + "\\" + fileName + "\\" + i + ".jpg", encoder, encParams);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            ImageInsetWord();
        }

        #region 根据pdf生成文档        
        private void ImageInsetWord()
        {
            //获取文件目录
            var projectDir = this.pdfSouceLb.Content.ToString();
            //获取所有的pdf文件
            var array = Directory.GetFiles(projectDir , "*.pdf");


            if (bool.Parse(moreRb.IsChecked.ToString()))
            {
                foreach (var file in array)
                {
                    //获取文件名称
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(file);

                    //创建word文档
                    string docPath = this.wordOutPutLb.Content+"\\" + fileName + ".docx"; // 替换为你的Word文档路径

                    InsertImagesToWord(projectDir +"\\"+ fileName, docPath);
                }
                MessageBox.Show("插入完成");
            }
            else
            {
                //创建word文档
                string docPath = this.wordOutPutLb.Content + "\\All.docx"; // 替换为你的Word文档路径
                InsertImagesToWord(projectDir, array, docPath);
            }           
        }
        void InsertImagesToWord(string projectDir, string[] folderPaths, string docPath)
        {
            try
            {
                // 创建Word应用程序对象
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true; // 可见，方便观察操作

                // 添加一个新的Word文档
                Word.Document doc = wordApp.Documents.Add();

                foreach (var folderPath in folderPaths)
                {
                    //获取文件名称
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(folderPath);

                    // 插入文字到Word文档
                    Word.Paragraph paragraph = doc.Paragraphs.Add(); // 创建一个新段落
                    Word.Range range = paragraph.Range;
                    range.Text = fileName;

                    // 获取图片文件夹中所有文件
                    string[] imageFiles = Directory.GetFiles(projectDir + "\\" + fileName, "*.jpg")
                        .OrderBy(f => Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(f))).ToArray(); // 根据实际图片格式修改扩展名s

                    foreach (string imagePath in imageFiles)
                    {
                        // 移动光标到下一行
                        paragraph.Range.InsertParagraphAfter();

                        // 获取当前段落的Range
                        Word.Range paragraphRange = paragraph.Range;

                        // 插入图片到Word文档
                        Word.InlineShape inlineShape = paragraphRange.InlineShapes.AddPicture(imagePath);

                        // 调整图片大小
                        // 设置宽度和高度，可以根据需要进行调整
                        inlineShape.Width = int.Parse(widthText.Text.ToString()); // 设置宽度，单位为磅（points）
                        inlineShape.Height = int.Parse(heightText.Text.ToString()); // 设置高度，单位为磅（points）                    
                    }
                }

                // 保存文档
                doc.SaveAs2(docPath);

                // 关闭Word应用程序
                wordApp.Quit();

                MessageBox.Show("插入完成");
            }
            catch (Exception ex)
            {

                MessageBox.Show("异常:" + ex.Message);
            }
        }




        void InsertImagesToWord(string folderPath, string docPath)
        {
            try
            {
                // 创建Word应用程序对象
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true; // 可见，方便观察操作

                // 添加一个新的Word文档
                Word.Document doc = wordApp.Documents.Add();

                // 插入文字到Word文档
                Word.Range range = doc.Range();
                range.Text = System.IO.Path.GetFileNameWithoutExtension(folderPath);

                // 获取图片文件夹中所有文件
                string[] imageFiles = Directory.GetFiles(folderPath, "*.jpg")
                    .OrderBy(f => Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(f))).ToArray(); // 根据实际图片格式修改扩展名s

                foreach (string imagePath in imageFiles)
                {
                    // 移动光标到下一行
                    wordApp.Selection.EndKey(Word.WdUnits.wdStory);

                    // 插入图片到Word文档
                    Word.InlineShape inlineShape = doc.InlineShapes.AddPicture(imagePath);
                    // 调整图片大小
                    // 设置宽度和高度，可以根据需要进行调整
                    inlineShape.Width = int.Parse(widthText.Text.ToString()); // 设置宽度，单位为磅（points）
                    inlineShape.Height = int.Parse(heightText.Text.ToString()); // 设置高度，单位为磅（points）               
                }



                // 保存文档
                doc.SaveAs2(docPath);

                // 关闭Word应用程序
                wordApp.Quit();
            }
            catch (Exception ex)
            {

                MessageBox.Show("异常:" + ex.Message);
            }
        
        }


        #endregion
        private void SelectPDF_Button_Click(object sender, RoutedEventArgs e)
        {
            // 打开文件夹对话框
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            pdfSouceLb.Content = dialog.SelectedPath.Trim();
        }

        private void OutPutWord_Button_Click(object sender, RoutedEventArgs e)
        {
            // 打开文件夹对话框
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            wordOutPutLb.Content = dialog.SelectedPath.Trim();
        }
    }
}
