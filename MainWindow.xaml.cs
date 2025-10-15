using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace H5PtoPPTX
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void BtnBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtInputFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtOutputFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private async void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            var inputFolder = TxtInputFolder.Text;
            var outputFolder = TxtOutputFolder.Text;

            if (string.IsNullOrEmpty(inputFolder) || !Directory.Exists(inputFolder))
            {
                System.Windows.MessageBox.Show("Please select a valid input folder.", "Invalid Input Folder", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                System.Windows.MessageBox.Show("Please select a valid output folder.", "Invalid Output Folder", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var h5pFiles = Directory.GetFiles(inputFolder, "*.h5p");

            if (h5pFiles.Length == 0)
            {
                System.Windows.MessageBox.Show("No .h5p files found in the selected input folder.", "No Files Found", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            LogListView.Items.Clear();
            ProgressBar.Value = 0;
            ProgressBar.Maximum = h5pFiles.Length;

            foreach (var h5pFilePath in h5pFiles)
            {
                var logItem = new LogItem { FileName = Path.GetFileName(h5pFilePath), Status = "Processing..." };
                LogListView.Items.Add(logItem);

                try
                {
                    await Task.Run(() => ProcessH5P(h5pFilePath, outputFolder));
                    logItem.Status = "Success";
                }
                catch (Exception ex)
                {
                    logItem.Status = "Error";
                    logItem.Message = ex.Message;
                }
                ProgressBar.Value++;
            }

            System.Windows.MessageBox.Show("Conversion complete!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ProcessH5P(string h5pFilePath, string outputFolder)
        {
            using (var archive = ZipFile.OpenRead(h5pFilePath))
            {
                var contentJsonEntry = archive.Entries.FirstOrDefault(entry => entry.FullName.Equals("content/content.json", StringComparison.OrdinalIgnoreCase));
                if (contentJsonEntry == null)
                {
                    throw new FileNotFoundException("Could not find 'content/content.json' in the H5P file.");
                }

                string contentJson;
                using (var stream = contentJsonEntry.Open())
                using (var reader = new StreamReader(stream))
                {
                    contentJson = reader.ReadToEnd();
                }

                var imagePaths = GetImagePaths(contentJson);

                if (imagePaths.Any())
                {
                    var pptxFilePath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(h5pFilePath) + ".pptx");
                    CreatePresentation(pptxFilePath, imagePaths, archive);
                }
            }
        }

        private List<string> GetImagePaths(string json)
        {
            var imagePaths = new List<string>();
            using (var document = JsonDocument.Parse(json))
            {
                var root = document.RootElement;
                if (root.TryGetProperty("presentation", out var presentation) && presentation.TryGetProperty("slides", out var slides))
                {
                    foreach (var slide in slides.EnumerateArray())
                    {
                        if (slide.TryGetProperty("elements", out var elements))
                        {
                            foreach (var element in elements.EnumerateArray())
                            {
                                if (element.TryGetProperty("action", out var action) &&
                                    action.TryGetProperty("library", out var library) &&
                                    library.GetString().StartsWith("H5P.Image", StringComparison.OrdinalIgnoreCase) &&
                                    action.TryGetProperty("params", out var parameters) &&
                                    parameters.TryGetProperty("file", out var file) &&
                                    file.TryGetProperty("path", out var path))
                                {
                                    imagePaths.Add(path.GetString());
                                }
                            }
                        }
                    }
                }
            }
            return imagePaths;
        }

        private void CreatePresentation(string pptxFilePath, List<string> imagePaths, ZipArchive archive)
        {
            using (var presentationDocument = PresentationDocument.Create(pptxFilePath, PresentationDocumentType.Presentation))
            {
                var presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new P.Presentation();

                CreatePresentationParts(presentationPart);

                uint slideIndex = 0;
                foreach (var imagePath in imagePaths)
                {
                    var imageEntry = archive.Entries.FirstOrDefault(entry => entry.FullName.Equals($"content/{imagePath}", StringComparison.OrdinalIgnoreCase));
                    if (imageEntry != null)
                    {
                        var slidePart = CreateSlidePart(presentationPart);
                        using (var stream = imageEntry.Open())
                        {
                            AddImageToSlide(slidePart, stream, slideIndex++, imagePath);
                        }
                        slidePart.Slide.Save();
                    }
                }

                presentationPart.Presentation.Save();
            }
        }

        private void CreatePresentationParts(PresentationPart presentationPart)
        {
            var slideMasterIdList = new P.SlideMasterIdList(new P.SlideMasterId() { Id = 2147483648U, RelationshipId = "rId1" });
            var slideIdList = new P.SlideIdList();
            var slideSize = new P.SlideSize() { Cx = 9144000, Cy = 5143500, Type = P.SlideSizeValues.Screen16x9 };
            var notesSize = new P.NotesSize() { Cx = 5143500, Cy = 9144000 };
            var defaultTextStyle = new P.DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList, slideIdList, slideSize, notesSize, defaultTextStyle);

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");

            slideMasterPart.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(new P.ShapeTree()),
                new P.SlideMasterIdList(new P.SlideMasterId() { Id = 2147483648U, RelationshipId = "rId1" }));

            slideLayoutPart.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(new P.ShapeTree()),
                new P.ColorMapOverride(new A.MasterColorMapping()),
                new P.ShapeTree());

            slideMasterPart.SlideMaster.Save();
            slideLayoutPart.SlideLayout.Save();
        }

        private SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree()));
            slidePart.AddPart(presentationPart.SlideMasterParts.First().SlideLayoutParts.First());
            var slideIdList = presentationPart.Presentation.SlideIdList;
            slideIdList.Append(new P.SlideId() { Id = (uint)(slideIdList.Count() + 256), RelationshipId = presentationPart.GetIdOfPart(slidePart) });
            return slidePart;
        }

        private void AddImageToSlide(SlidePart slidePart, Stream imageStream, uint slideIndex, string imagePath)
        {
            var imagePartType = GetImagePartType(Path.GetExtension(imagePath));
            var imagePart = slidePart.AddImagePart(imagePartType, $"rId{slideIndex + 2}");
            imagePart.FeedData(imageStream);

            var tree = slidePart.Slide.CommonSlideData.ShapeTree;

            var picture = new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties() { Id = (uint)tree.Count() + 1, Name = $"Picture {slideIndex + 1}" },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(
                    new A.Blip() { Embed = slidePart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 0, Y = 0 },
                        new A.Extents() { Cx = 9144000, Cy = 5143500 }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
            );

            tree.Append(picture);
        }

        private string GetImagePartType(string extension)
        {
            switch (extension.ToLower())
            {
                case ".png":
                    return "image/png";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".tiff":
                    return "image/tiff";
                default:
                    return "image/jpeg";
            }
        }
    }

    public class LogItem : INotifyPropertyChanged
    {
        private string _fileName;
        public string FileName
        {
            get { return _fileName; }
            set
            {
                _fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }



        private string _status;
        public string Status
        {
            get { return _status; }
            set
            {
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        private string _message;
        public string Message
        {
            get { return _message; }
            set
            {
                _message = value;
                OnPropertyChanged(nameof(Message));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}