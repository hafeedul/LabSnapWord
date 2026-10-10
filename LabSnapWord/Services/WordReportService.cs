using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace LabSnapWord.Services
{
    public static class WordReportService
    {
        private const long EmuPerInch = 914400L;
        private const long EmuPerPixelAt96Dpi = 9525L;
        private const double MaxWidthInches = 6.2;
        private const double MaxHeightInches = 7.2;

        public static void EnsureOrCreateDocument(string filePath, string? reportTitle = null, bool overwrite = false)
        {
            if (File.Exists(filePath) && !overwrite)
            {
                return;
            }

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    Body body = new Body();

                    if (!string.IsNullOrWhiteSpace(reportTitle))
                    {
                        Paragraph titlePara = CreateStyledParagraph(
                            reportTitle.Trim(),
                            fontSizeHalfPoints: "36",
                            hexColor: "1E293B",
                            bold: true,
                            spaceBeforeTwips: "120",
                            spaceAfterTwips: "240");
                        body.AppendChild(titlePara);
                    }

                    mainPart.Document = new Document(body);
                    mainPart.Document.Save();
                }
            }
            catch (IOException)
            {
            }
        }

        public static bool AppendScreenshotEntry(
            string filePath,
            byte[] pngBytes,
            int pixelWidth,
            int pixelHeight,
            string? taskHeading,
            string? notes)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Please choose a target Word (.docx) file first.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                EnsureOrCreateDocument(filePath);
            }

            try
            {
                AppendViaOpenXml(filePath, pngBytes, pixelWidth, pixelHeight, taskHeading, notes);
                return false;
            }
            catch (IOException)
            {
                AppendViaRunningWordCom(filePath, pngBytes, taskHeading, notes);
                return true;
            }
        }

        private static void AppendViaOpenXml(
            string filePath,
            byte[] pngBytes,
            int pixelWidth,
            int pixelHeight,
            string? taskHeading,
            string? notes)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, isEditable: true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
                if (mainPart.Document == null)
                {
                    mainPart.Document = new Document(new Body());
                }

                Body body = mainPart.Document.Body ?? mainPart.Document.AppendChild(new Body());

                if (!string.IsNullOrWhiteSpace(taskHeading))
                {
                    Paragraph headingPara = CreateStyledParagraph(
                        taskHeading.Trim(),
                        fontSizeHalfPoints: "28",
                        hexColor: "0F172A",
                        bold: true,
                        spaceBeforeTwips: "240",
                        spaceAfterTwips: "80");
                    body.AppendChild(headingPara);
                }

                if (!string.IsNullOrWhiteSpace(notes))
                {
                    Paragraph notesPara = CreateStyledParagraph(
                        notes.Trim(),
                        fontSizeHalfPoints: "22",
                        hexColor: "334155",
                        bold: false,
                        spaceBeforeTwips: "0",
                        spaceAfterTwips: "100");
                    body.AppendChild(notesPara);
                }

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (MemoryStream ms = new MemoryStream(pngBytes))
                {
                    imagePart.FeedData(ms);
                }

                string relationshipId = mainPart.GetIdOfPart(imagePart);
                uint nextDocPrId = GetNextDocPropertiesId(body);

                CalculateImageEmuSize(pixelWidth, pixelHeight, out long widthEmu, out long heightEmu);

                Drawing drawing = BuildImageDrawing(relationshipId, nextDocPrId, widthEmu, heightEmu);

                Paragraph imagePara = new Paragraph(
                    new ParagraphProperties(
                        new SpacingBetweenLines { Before = "60", After = "240" },
                        new Justification { Val = JustificationValues.Left }),
                    new Run(drawing));

                body.AppendChild(imagePara);
                mainPart.Document.Save();
            }
        }

        private static void AppendViaRunningWordCom(
            string filePath,
            byte[] pngBytes,
            string? taskHeading,
            string? notes)
        {
            string tempPngPath = Path.Combine(Path.GetTempPath(), $"LabSnap_{Guid.NewGuid():N}.png");
            File.WriteAllBytes(tempPngPath, pngBytes);

            try
            {
                string psScript = @"
$ErrorActionPreference = 'Stop'
$docPath = [System.IO.Path]::GetFullPath($env:LABSNAP_DOC_PATH)
$imgPath = $env:LABSNAP_IMG_PATH
$heading = $env:LABSNAP_HEADING
$notes   = $env:LABSNAP_NOTES
$maxW    = 446.0
$maxH    = 518.0

$word = $null
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
}

$targetDoc = $null
foreach ($d in $word.Documents) {
    try {
        if ([string]::Equals([System.IO.Path]::GetFullPath($d.FullName), $docPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $targetDoc = $d
            break
        }
    } catch {}
}

if ($null -eq $targetDoc) {
    $targetDoc = $word.Documents.Open($docPath)
}

$range = $targetDoc.Content
$range.Collapse(0)
if ($targetDoc.Content.End -gt 1) {
    $range.InsertParagraphAfter()
    $range.Collapse(0)
}

if (-not [string]::IsNullOrWhiteSpace($heading)) {
    $range.Text = $heading
    $range.Font.Name = 'Calibri'
    $range.Font.Size = 14
    $range.Font.Bold = 1
    $range.ParagraphFormat.SpaceBefore = 12
    $range.ParagraphFormat.SpaceAfter = 4
    $range.InsertParagraphAfter()
    $range.Collapse(0)
}

if (-not [string]::IsNullOrWhiteSpace($notes)) {
    $range.Text = $notes
    $range.Font.Name = 'Calibri'
    $range.Font.Size = 11
    $range.Font.Bold = 0
    $range.ParagraphFormat.SpaceBefore = 0
    $range.ParagraphFormat.SpaceAfter = 6
    $range.InsertParagraphAfter()
    $range.Collapse(0)
}

$range.Font.Bold = 0
$range.ParagraphFormat.SpaceBefore = 3
$range.ParagraphFormat.SpaceAfter = 12
$shape = $targetDoc.InlineShapes.AddPicture($imgPath, $false, $true, $range)
$shape.LockAspectRatio = -1

if ($shape.Width -gt $maxW) {
    $scale = $maxW / [double]$shape.Width
    $shape.Width = [single]$maxW
    $shape.Height = [single]([double]$shape.Height * $scale)
}
if ($shape.Height -gt $maxH) {
    $scale = $maxH / [double]$shape.Height
    $shape.Height = [single]$maxH
    $shape.Width = [single]([double]$shape.Width * $scale)
}

$endRange = $targetDoc.Content
$endRange.Collapse(0)
$endRange.InsertParagraphAfter()

$targetDoc.Save()
";

                string encodedCommand = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(psScript));

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedCommand}",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                psi.Environment["LABSNAP_DOC_PATH"] = Path.GetFullPath(filePath);
                psi.Environment["LABSNAP_IMG_PATH"] = tempPngPath;
                psi.Environment["LABSNAP_HEADING"] = taskHeading ?? string.Empty;
                psi.Environment["LABSNAP_NOTES"] = notes ?? string.Empty;

                using (Process? proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        throw new InvalidOperationException("Could not start Word automation bridge.");
                    }

                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit(15000);

                    if (proc.ExitCode != 0)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(stderr)
                                ? "Microsoft Word is busy (for example, a dialog box is open in Word). Please close any open Word dialog box and try again."
                                : stderr.Trim());
                    }
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPngPath))
                    {
                        File.Delete(tempPngPath);
                    }
                }
                catch
                {
                }
            }
        }

        private static Paragraph CreateStyledParagraph(
            string text,
            string fontSizeHalfPoints,
            string hexColor,
            bool bold,
            string spaceBeforeTwips,
            string spaceAfterTwips)
        {
            RunProperties runProps = new RunProperties(
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" },
                new Color { Val = hexColor },
                new FontSize { Val = fontSizeHalfPoints },
                new FontSizeComplexScript { Val = fontSizeHalfPoints });

            if (bold)
            {
                runProps.AppendChild(new Bold());
            }

            Run run = new Run(runProps);
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    run.AppendChild(new Break());
                }
                run.AppendChild(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            }

            ParagraphProperties paraProps = new ParagraphProperties(
                new SpacingBetweenLines { Before = spaceBeforeTwips, After = spaceAfterTwips });

            return new Paragraph(paraProps, run);
        }

        private static void CalculateImageEmuSize(int pixelWidth, int pixelHeight, out long widthEmu, out long heightEmu)
        {
            widthEmu = Math.Max(1L, pixelWidth) * EmuPerPixelAt96Dpi;
            heightEmu = Math.Max(1L, pixelHeight) * EmuPerPixelAt96Dpi;

            long maxWidthEmu = (long)(MaxWidthInches * EmuPerInch);
            long maxHeightEmu = (long)(MaxHeightInches * EmuPerInch);

            if (widthEmu > maxWidthEmu)
            {
                double scale = (double)maxWidthEmu / widthEmu;
                widthEmu = maxWidthEmu;
                heightEmu = (long)(heightEmu * scale);
            }

            if (heightEmu > maxHeightEmu)
            {
                double scale = (double)maxHeightEmu / heightEmu;
                heightEmu = maxHeightEmu;
                widthEmu = (long)(widthEmu * scale);
            }
        }

        private static uint GetNextDocPropertiesId(Body body)
        {
            uint maxId = 0;
            foreach (DW.DocProperties prop in body.Descendants<DW.DocProperties>())
            {
                if (prop.Id is { HasValue: true } idVal && idVal.Value > maxId)
                {
                    maxId = idVal.Value;
                }
            }
            return maxId + 1;
        }

        private static Drawing BuildImageDrawing(string relationshipId, uint docPrId, long widthEmu, long heightEmu)
        {
            string imageName = $"LabScreenshot_{docPrId}.png";

            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = docPrId, Name = imageName },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = docPrId, Name = imageName },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    })))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });
        }
    }
}
