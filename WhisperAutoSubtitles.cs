using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        try
        {
            var config = LoadConfig(vegas);
            if (config == null)
            {
                return;
            }

            EnsureWhisperExecutable(config.WhisperExePath);

            string tempOutputDir = Path.Combine(Path.GetTempPath(), "VegasWhisperSubs", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempOutputDir);

            string jsonPath = RunWhisperWithProgress(config, tempOutputDir);
            var captions = ParseWhisperJson(jsonPath, config.SplitText ? config.WordsPerCaption : 0);
            if (captions.Count == 0)
            {
                throw new ApplicationException("Whisper finished, but no subtitle entries were found in the output.");
            }

            AddSubtitlesToTimeline(vegas, captions, config);

            MessageBox.Show(
                "Subtitles created successfully.\nEntries: " + captions.Count,
                "Whisper Subtitles",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Whisper Subtitle Script Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private class SubtitleConfig
    {
        public string InputMediaPath;
        public string WhisperExePath;
        public string Model;
        public string Language;
        public string FontName;
        public int FontSize;
        public Color TextColor;
        public Color OutlineColor;
        public int OutlineWidth;
        public bool SplitText;
        public int WordsPerCaption;
        public bool UseLineBreaks;
        public int WordsPerLine;
    }

    private class Caption
    {
        public TimeSpan Start;
        public TimeSpan End;
        public string Text;
    }

    private SubtitleConfig LoadConfig(Vegas vegas)
    {
        string mediaPath = TryGetSelectedMediaPath(vegas);
        if (string.IsNullOrWhiteSpace(mediaPath) || !File.Exists(mediaPath))
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Pick media/audio file to transcribe";
                ofd.Filter = "Media files|*.wav;*.mp3;*.m4a;*.flac;*.aac;*.ogg;*.mp4;*.mov;*.mkv;*.avi|All files|*.*";
                ofd.Multiselect = false;
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    return null;
                }

                mediaPath = ofd.FileName;
            }
        }

        string whisperExe = Environment.GetEnvironmentVariable("WHISPER_EXE");
        if (string.IsNullOrWhiteSpace(whisperExe))
        {
            whisperExe = "whisper";
        }

        string model = Prompt("Whisper model", "small", "Whisper model to use (tiny, small, medium, large):");
        if (model == null)
        {
            return null;
        }

        string language = Prompt("Whisper language", "en", "Language code (example: en, de, fr). Leave empty for auto-detect:");
        if (language == null)
        {
            return null;
        }

        string exeInput = Prompt(
            "Whisper executable",
            whisperExe,
            "Whisper executable path or command name.\nSet WHISPER_EXE env var to avoid this prompt next time:"
        );
        if (exeInput == null)
        {
            return null;
        }

        // Ask about text splitting mode
        string splitModeInput = Prompt(
            "Text Split Mode",
            "yes",
            "Split text into short captions?\n\n" +
            "• yes - Split into smaller chunks (default: 4 words each)\n" +
            "• no - Keep original sentence/segment timing\n" +
            "• [number] - Split with custom words per caption (e.g., 3, 5, 6)"
        );
        if (splitModeInput == null)
        {
            return null;
        }

        bool splitText = true;
        int wordsPerCaption = 4;
        splitModeInput = splitModeInput.Trim().ToLowerInvariant();
        
        if (splitModeInput == "no" || splitModeInput == "n" || splitModeInput == "false")
        {
            splitText = false;
        }
        else if (splitModeInput != "yes" && splitModeInput != "y" && splitModeInput != "true")
        {
            int parsed;
            if (int.TryParse(splitModeInput, out parsed) && parsed > 0)
            {
                wordsPerCaption = parsed;
            }
        }

        // Ask about line breaks within captions
        string lineBreakInput = Prompt(
            "Line Breaks",
            "no",
            "Add line breaks within captions?\n\n" +
            "• no - No line breaks (single line captions)\n" +
            "• yes - Add line breaks (default: every 4 words)\n" +
            "• [number] - Add line break after X words (e.g., 3, 5, 6)"
        );
        if (lineBreakInput == null)
        {
            return null;
        }

        bool useLineBreaks = false;
        int wordsPerLine = 4;
        lineBreakInput = lineBreakInput.Trim().ToLowerInvariant();
        
        if (lineBreakInput == "yes" || lineBreakInput == "y" || lineBreakInput == "true")
        {
            useLineBreaks = true;
        }
        else if (lineBreakInput != "no" && lineBreakInput != "n" && lineBreakInput != "false")
        {
            int parsed;
            if (int.TryParse(lineBreakInput, out parsed) && parsed > 0)
            {
                useLineBreaks = true;
                wordsPerLine = parsed;
            }
        }

        // Show style settings dialog with preview
        SubtitleStyleSettings styleSettings = ShowStyleDialog();
        if (styleSettings == null)
        {
            return null;
        }

        return new SubtitleConfig
        {
            InputMediaPath = mediaPath,
            WhisperExePath = ResolveExecutablePath(exeInput),
            Model = model.Trim(),
            Language = language.Trim(),
            FontName = styleSettings.FontName,
            FontSize = styleSettings.FontSize,
            TextColor = styleSettings.TextColor,
            OutlineColor = styleSettings.OutlineColor,
            OutlineWidth = styleSettings.OutlineWidth,
            SplitText = splitText,
            WordsPerCaption = wordsPerCaption,
            UseLineBreaks = useLineBreaks,
            WordsPerLine = wordsPerLine
        };
    }

    private class SubtitleStyleSettings
    {
        public string FontName;
        public int FontSize;
        public Color TextColor;
        public Color OutlineColor;
        public int OutlineWidth;
    }

    private SubtitleStyleSettings ShowStyleDialog()
    {
        string defaultFont = Environment.GetEnvironmentVariable("WHISPER_SUB_FONT");
        if (string.IsNullOrWhiteSpace(defaultFont))
        {
            defaultFont = "Arial";
        }

        SubtitleStyleSettings settings = new SubtitleStyleSettings
        {
            FontName = defaultFont,
            FontSize = 24,
            TextColor = Color.White,
            OutlineColor = Color.Black,
            OutlineWidth = 2
        };

        using (Form form = new Form())
        {
            form.Text = "Subtitle Style Settings";
            form.Width = 500;
            form.Height = 480;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;

            int y = 15;

            // Font selection
            Label fontLabel = new Label { Text = "Font:", Left = 15, Top = y, Width = 80 };
            ComboBox fontCombo = new ComboBox { Left = 100, Top = y - 3, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (FontFamily ff in FontFamily.Families)
            {
                fontCombo.Items.Add(ff.Name);
            }
            fontCombo.SelectedItem = defaultFont;
            if (fontCombo.SelectedIndex < 0 && fontCombo.Items.Count > 0)
            {
                fontCombo.SelectedIndex = 0;
            }

            y += 35;

            // Font size
            Label sizeLabel = new Label { Text = "Font Size:", Left = 15, Top = y, Width = 80 };
            NumericUpDown sizeInput = new NumericUpDown { Left = 100, Top = y - 3, Width = 80, Minimum = 8, Maximum = 120, Value = 24 };
            Label sizePreviewLabel = new Label { Text = "pt", Left = 185, Top = y, Width = 30 };

            y += 35;

            // Text color
            Label textColorLabel = new Label { Text = "Text Color:", Left = 15, Top = y, Width = 80 };
            Panel textColorPanel = new Panel { Left = 100, Top = y - 3, Width = 40, Height = 25, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            Button textColorBtn = new Button { Text = "Pick...", Left = 150, Top = y - 4, Width = 60, Height = 27 };

            y += 35;

            // Outline color
            Label outlineColorLabel = new Label { Text = "Outline Color:", Left = 15, Top = y, Width = 80 };
            Panel outlineColorPanel = new Panel { Left = 100, Top = y - 3, Width = 40, Height = 25, BackColor = Color.Black, BorderStyle = BorderStyle.FixedSingle };
            Button outlineColorBtn = new Button { Text = "Pick...", Left = 150, Top = y - 4, Width = 60, Height = 27 };

            y += 35;

            // Outline width
            Label outlineWidthLabel = new Label { Text = "Outline Width:", Left = 15, Top = y, Width = 80 };
            NumericUpDown outlineWidthInput = new NumericUpDown { Left = 100, Top = y - 3, Width = 80, Minimum = 0, Maximum = 10, Value = 2 };
            Label outlineWidthPx = new Label { Text = "px", Left = 185, Top = y, Width = 30 };

            y += 40;

            // Preview panel
            Label previewLabel = new Label { Text = "Preview:", Left = 15, Top = y, Width = 80 };
            y += 20;
            Panel previewPanel = new Panel { Left = 15, Top = y, Width = 455, Height = 120, BackColor = Color.FromArgb(40, 40, 40), BorderStyle = BorderStyle.FixedSingle };

            y += 135;

            // Buttons
            Button okButton = new Button { Text = "OK", Left = 295, Top = y, Width = 80, DialogResult = DialogResult.OK };
            Button cancelButton = new Button { Text = "Cancel", Left = 385, Top = y, Width = 80, DialogResult = DialogResult.Cancel };

            // Preview paint handler
            previewPanel.Paint += (sender, e) =>
            {
                string sampleText = "Sample Subtitle";
                string fontName = fontCombo.SelectedItem != null ? fontCombo.SelectedItem.ToString() : "Arial";
                int fontSize = (int)sizeInput.Value;
                Color textColor = textColorPanel.BackColor;
                Color outlineColor = outlineColorPanel.BackColor;
                int outlineWidth = (int)outlineWidthInput.Value;

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                using (Font font = new Font(fontName, fontSize, FontStyle.Bold, GraphicsUnit.Point))
                using (GraphicsPath path = new GraphicsPath())
                {
                    SizeF textSize = e.Graphics.MeasureString(sampleText, font);
                    float x = (previewPanel.Width - textSize.Width) / 2;
                    float y2 = (previewPanel.Height - textSize.Height) / 2;

                    path.AddString(sampleText, font.FontFamily, (int)FontStyle.Bold, fontSize * 1.33f, new PointF(x, y2), StringFormat.GenericDefault);

                    if (outlineWidth > 0)
                    {
                        // Scale down to better match Vegas output (Vegas uses thinner strokes)
                        float scaledOutline = outlineWidth * 0.75f;
                        using (Pen outlinePen = new Pen(outlineColor, scaledOutline))
                        {
                            outlinePen.LineJoin = LineJoin.Round;
                            e.Graphics.DrawPath(outlinePen, path);
                        }
                    }

                    using (SolidBrush textBrush = new SolidBrush(textColor))
                    {
                        e.Graphics.FillPath(textBrush, path);
                    }
                }
            };

            // Update preview when settings change
            EventHandler updatePreview = (sender, e) => previewPanel.Invalidate();
            fontCombo.SelectedIndexChanged += updatePreview;
            sizeInput.ValueChanged += updatePreview;
            outlineWidthInput.ValueChanged += updatePreview;

            // Color picker handlers
            textColorBtn.Click += (sender, e) =>
            {
                using (ColorDialog cd = new ColorDialog { Color = textColorPanel.BackColor, FullOpen = true })
                {
                    if (cd.ShowDialog() == DialogResult.OK)
                    {
                        textColorPanel.BackColor = cd.Color;
                        previewPanel.Invalidate();
                    }
                }
            };

            outlineColorBtn.Click += (sender, e) =>
            {
                using (ColorDialog cd = new ColorDialog { Color = outlineColorPanel.BackColor, FullOpen = true })
                {
                    if (cd.ShowDialog() == DialogResult.OK)
                    {
                        outlineColorPanel.BackColor = cd.Color;
                        previewPanel.Invalidate();
                    }
                }
            };

            // Add controls
            form.Controls.AddRange(new Control[] {
                fontLabel, fontCombo,
                sizeLabel, sizeInput, sizePreviewLabel,
                textColorLabel, textColorPanel, textColorBtn,
                outlineColorLabel, outlineColorPanel, outlineColorBtn,
                outlineWidthLabel, outlineWidthInput, outlineWidthPx,
                previewLabel, previewPanel,
                okButton, cancelButton
            });

            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            if (form.ShowDialog() == DialogResult.OK)
            {
                settings.FontName = fontCombo.SelectedItem != null ? fontCombo.SelectedItem.ToString() : "Arial";
                settings.FontSize = (int)sizeInput.Value;
                settings.TextColor = textColorPanel.BackColor;
                settings.OutlineColor = outlineColorPanel.BackColor;
                settings.OutlineWidth = (int)outlineWidthInput.Value;
                return settings;
            }

            return null;
        }
    }

    private static string Prompt(string title, string defaultValue, string message)
    {
        using (Form form = new Form())
        using (Label label = new Label())
        using (TextBox textBox = new TextBox())
        using (Button okButton = new Button())
        using (Button cancelButton = new Button())
        {
            form.Text = title;
            form.Width = 720;
            form.Height = 170;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;

            label.Text = message;
            label.Left = 10;
            label.Top = 10;
            label.Width = 680;
            label.Height = 36;

            textBox.Left = 10;
            textBox.Top = 55;
            textBox.Width = 680;
            textBox.Text = defaultValue;

            okButton.Text = "OK";
            okButton.Left = 525;
            okButton.Top = 90;
            okButton.Width = 80;
            okButton.DialogResult = DialogResult.OK;

            cancelButton.Text = "Cancel";
            cancelButton.Left = 610;
            cancelButton.Top = 90;
            cancelButton.Width = 80;
            cancelButton.DialogResult = DialogResult.Cancel;

            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(okButton);
            form.Controls.Add(cancelButton);
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
        }
    }

    private static string TryGetSelectedMediaPath(Vegas vegas)
    {
        // First try to get media from a selected event
        foreach (Track track in vegas.Project.Tracks)
        {
            foreach (TrackEvent ev in track.Events)
            {
                if (!ev.Selected || ev.ActiveTake == null || ev.ActiveTake.Media == null)
                {
                    continue;
                }

                string path = ev.ActiveTake.Media.FilePath;
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    return path;
                }
            }
        }

        // Fallback: get the first media with audio from any track in the timeline
        foreach (Track track in vegas.Project.Tracks)
        {
            foreach (TrackEvent ev in track.Events)
            {
                if (ev.ActiveTake == null || ev.ActiveTake.Media == null)
                {
                    continue;
                }

                string path = ev.ActiveTake.Media.FilePath;
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }

    private static void EnsureWhisperExecutable(string whisperExe)
    {
        if (string.IsNullOrWhiteSpace(whisperExe))
        {
            throw new ApplicationException("Whisper executable path is empty.");
        }

        if (!File.Exists(whisperExe))
        {
            throw new FileNotFoundException(
                "Whisper executable was not found.\n" +
                "Checked: " + whisperExe + "\n" +
                "Set WHISPER_EXE to a full path, e.g. C:\\Users\\<you>\\AppData\\Roaming\\Python\\Python3xx\\Scripts\\whisper.exe",
                whisperExe
            );
        }
    }

    private static string RunWhisper(SubtitleConfig config, string outputDir)
    {
        string outputBase = Path.GetFileNameWithoutExtension(config.InputMediaPath);
        string args = BuildWhisperArgs(config, outputDir);

        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = config.WhisperExePath,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = outputDir,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        string stdout;
        string stderr;
        int exitCode;

        Process process;
        try
        {
            process = Process.Start(psi);
        }
        catch (Win32Exception ex)
        {
            throw new ApplicationException(
                "Could not start Whisper executable.\n" +
                "Executable: " + config.WhisperExePath + "\n" +
                "Details: " + ex.Message
            );
        }

        using (process)
        {
            if (process == null)
            {
                throw new ApplicationException("Failed to start Whisper process.");
            }

            stdout = process.StandardOutput.ReadToEnd();
            stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            exitCode = process.ExitCode;
        }

        string expectedJson = Path.Combine(outputDir, outputBase + ".json");
        if (File.Exists(expectedJson))
        {
            return expectedJson;
        }

        string[] allJsons = Directory.GetFiles(outputDir, "*.json", SearchOption.TopDirectoryOnly);
        if (allJsons.Length > 0)
        {
            string newest = allJsons[0];
            DateTime newestTime = File.GetLastWriteTimeUtc(newest);
            for (int i = 1; i < allJsons.Length; i++)
            {
                DateTime candidateTime = File.GetLastWriteTimeUtc(allJsons[i]);
                if (candidateTime > newestTime)
                {
                    newest = allJsons[i];
                    newestTime = candidateTime;
                }
            }

            return newest;
        }

        string errorText = "Whisper did not produce a JSON file.";
        if (exitCode != 0)
        {
            errorText += "\nExit code: " + exitCode;
        }

        if (!string.IsNullOrWhiteSpace(stderr))
        {
            string lower = stderr.ToLowerInvariant();
            if (lower.Contains("ffmpeg") && (lower.Contains("not found") || lower.Contains("no such file")))
            {
                errorText += "\nffmpeg appears to be missing. Install ffmpeg and add its bin folder to PATH, e.g. C:\\ffmpeg\\bin";
            }

            errorText += "\nStderr:\n" + stderr;
        }
        else if (!string.IsNullOrWhiteSpace(stdout))
        {
            errorText += "\nStdout:\n" + stdout;
        }

        throw new ApplicationException(errorText);
    }

    private static string RunWhisperWithProgress(SubtitleConfig config, string outputDir)
    {
        string srtPath = null;
        Exception workerError = null;

        using (Form progressForm = new Form())
        using (Label statusLabel = new Label())
        using (ProgressBar progressBar = new ProgressBar())
        using (System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer())
        using (BackgroundWorker worker = new BackgroundWorker())
        {
            progressForm.Text = "Whisper Transcribing";
            progressForm.Width = 560;
            progressForm.Height = 170;
            progressForm.StartPosition = FormStartPosition.CenterScreen;
            progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            progressForm.MaximizeBox = false;
            progressForm.MinimizeBox = false;
            progressForm.ControlBox = false;

            statusLabel.Left = 12;
            statusLabel.Top = 12;
            statusLabel.Width = 520;
            statusLabel.Height = 40;
            statusLabel.Text = "Running Whisper... this can take a while on larger files.";

            progressBar.Left = 12;
            progressBar.Top = 62;
            progressBar.Width = 520;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;

            progressForm.Controls.Add(statusLabel);
            progressForm.Controls.Add(progressBar);

            DateTime startedAt = DateTime.UtcNow;
            timer.Interval = 500;
            timer.Tick += delegate
            {
                TimeSpan elapsed = DateTime.UtcNow - startedAt;
                statusLabel.Text = "Running Whisper... Elapsed: " + elapsed.ToString(@"hh\:mm\:ss");
            };
            timer.Start();

            worker.DoWork += delegate
            {
                srtPath = RunWhisper(config, outputDir);
            };

            worker.RunWorkerCompleted += delegate (object sender, RunWorkerCompletedEventArgs e)
            {
                timer.Stop();
                if (e.Error != null)
                {
                    workerError = e.Error;
                }

                progressForm.Close();
            };

            worker.RunWorkerAsync();
            progressForm.ShowDialog();
        }

        if (workerError != null)
        {
            throw workerError;
        }

        if (string.IsNullOrWhiteSpace(srtPath))
        {
            throw new ApplicationException("Whisper finished without returning a subtitle path.");
        }

        return srtPath;
    }

    private static string ResolveExecutablePath(string commandOrPath)
    {
        if (string.IsNullOrWhiteSpace(commandOrPath))
        {
            return commandOrPath;
        }

        string trimmed = commandOrPath.Trim();
        if (File.Exists(trimmed))
        {
            return Path.GetFullPath(trimmed);
        }

        if (trimmed.IndexOf(Path.DirectorySeparatorChar) >= 0 || trimmed.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
        {
            string candidate = Path.GetFullPath(trimmed);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        string[] pathExts = (Environment.GetEnvironmentVariable("PATHEXT") ?? ".EXE;.CMD;.BAT")
            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        string[] pathDirs = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

        bool hasExt = Path.HasExtension(trimmed);
        for (int i = 0; i < pathDirs.Length; i++)
        {
            string dir = pathDirs[i].Trim();
            if (dir.Length == 0)
            {
                continue;
            }

            if (hasExt)
            {
                string candidate = Path.Combine(dir, trimmed);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            else
            {
                for (int e = 0; e < pathExts.Length; e++)
                {
                    string candidate = Path.Combine(dir, trimmed + pathExts[e]);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }
        }

        return trimmed;
    }

    private static string BuildWhisperArgs(SubtitleConfig config, string outputDir)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(Quote(config.InputMediaPath));
        sb.Append(" --model ").Append(config.Model);
        sb.Append(" --output_format json");
        sb.Append(" --output_dir ").Append(Quote(outputDir));
        sb.Append(" --word_timestamps True");
        sb.Append(" --verbose False --task transcribe");

        if (!string.IsNullOrWhiteSpace(config.Language))
        {
            sb.Append(" --language ").Append(config.Language);
        }

        return sb.ToString();
    }

    private static string Quote(string value)
    {
        if (value == null)
        {
            return "\"\"";
        }

        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }

    private static List<Caption> ParseSrt(string path)
    {
        string raw = File.ReadAllText(path, Encoding.UTF8);
        raw = raw.Replace("\r\n", "\n").Replace("\r", "\n");

        string[] blocks = Regex.Split(raw.Trim(), @"\n\s*\n");
        List<Caption> captions = new List<Caption>();

        foreach (string block in blocks)
        {
            string[] rawLines = block.Split(new[] { '\n' }, StringSplitOptions.None);
            List<string> keptLines = new List<string>();
            for (int i = 0; i < rawLines.Length; i++)
            {
                string trimmed = rawLines[i].TrimEnd();
                if (trimmed.Length > 0)
                {
                    keptLines.Add(trimmed);
                }
            }

            string[] lines = keptLines.ToArray();
            if (lines.Length < 2)
            {
                continue;
            }

            int timeLineIndex = lines[0].Contains("-->") ? 0 : 1;
            if (timeLineIndex >= lines.Length)
            {
                continue;
            }

            string timeLine = lines[timeLineIndex];
            string[] parts = timeLine.Split(new[] { "-->" }, StringSplitOptions.None);
            if (parts.Length != 2)
            {
                continue;
            }

            TimeSpan start;
            TimeSpan end;
            if (!TryParseSrtTimestamp(parts[0].Trim(), out start) || !TryParseSrtTimestamp(parts[1].Trim(), out end))
            {
                continue;
            }

            StringBuilder textBuilder = new StringBuilder();
            for (int i = timeLineIndex + 1; i < lines.Length; i++)
            {
                if (textBuilder.Length > 0)
                {
                    textBuilder.Append("\n");
                }

                textBuilder.Append(lines[i]);
            }

            string text = textBuilder.ToString().Trim();
            if (text.Length == 0)
            {
                continue;
            }

            if (end <= start)
            {
                end = start + TimeSpan.FromMilliseconds(500);
            }

            captions.Add(new Caption { Start = start, End = end, Text = text });
        }

        return captions;
    }

    private class WordInfo
    {
        public string Word;
        public double Start;
        public double End;
    }

    private static List<Caption> ParseWhisperJson(string path, int wordsPerCaption)
    {
        string json = File.ReadAllText(path, Encoding.UTF8);
        
        // If wordsPerCaption is 0 or less, parse segment-level captions (no splitting)
        if (wordsPerCaption <= 0)
        {
            return ParseWhisperJsonSegments(json);
        }
        
        List<WordInfo> allWords = new List<WordInfo>();

        // Parse segments array from Whisper JSON
        // Format: {"segments": [{"words": [{"word": "Hello", "start": 0.0, "end": 0.5}, ...]}]}
        MatchCollection segmentMatches = Regex.Matches(json, @"""words""\s*:\s*\[(.*?)\]", RegexOptions.Singleline);
        
        foreach (Match segmentMatch in segmentMatches)
        {
            string wordsArrayContent = segmentMatch.Groups[1].Value;
            
            // Match each word object
            MatchCollection wordMatches = Regex.Matches(wordsArrayContent, 
                @"\{\s*""word""\s*:\s*""([^""]*)""\s*,\s*""start""\s*:\s*([\d.]+)\s*,\s*""end""\s*:\s*([\d.]+)", 
                RegexOptions.Singleline);
            
            foreach (Match wordMatch in wordMatches)
            {
                string word = wordMatch.Groups[1].Value;
                double start, end;
                
                if (double.TryParse(wordMatch.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out start) &&
                    double.TryParse(wordMatch.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out end))
                {
                    // Decode Unicode escapes and clean up the word
                    word = DecodeJsonUnicode(word).Trim();
                    if (!string.IsNullOrEmpty(word))
                    {
                        allWords.Add(new WordInfo { Word = word, Start = start, End = end });
                    }
                }
            }
        }

        // Group words into captions with accurate timestamps
        List<Caption> captions = new List<Caption>();
        
        for (int i = 0; i < allWords.Count; i += wordsPerCaption)
        {
            int count = Math.Min(wordsPerCaption, allWords.Count - i);
            
            StringBuilder text = new StringBuilder();
            double captionStart = allWords[i].Start;
            double captionEnd = allWords[i + count - 1].End;
            
            for (int j = 0; j < count; j++)
            {
                if (j > 0)
                {
                    text.Append(" ");
                }
                text.Append(allWords[i + j].Word);
            }
            
            if (captionEnd <= captionStart)
            {
                captionEnd = captionStart + 0.3;
            }
            
            captions.Add(new Caption
            {
                Start = TimeSpan.FromSeconds(captionStart),
                End = TimeSpan.FromSeconds(captionEnd),
                Text = text.ToString()
            });
        }

        return captions;
    }

    private static List<Caption> ParseWhisperJsonSegments(string json)
    {
        List<Caption> captions = new List<Caption>();
        
        // Parse segments array from Whisper JSON
        // Format: {"segments": [{"start": 0.0, "end": 2.5, "text": "Hello world", ...}, ...]}
        MatchCollection segmentMatches = Regex.Matches(json, 
            @"\{\s*[^{}]*""start""\s*:\s*([\d.]+)\s*,\s*""end""\s*:\s*([\d.]+)\s*,\s*""text""\s*:\s*""([^""]*)""",
            RegexOptions.Singleline);
        
        foreach (Match match in segmentMatches)
        {
            double start, end;
            
            if (double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out start) &&
                double.TryParse(match.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out end))
            {
                string text = DecodeJsonUnicode(match.Groups[3].Value).Trim();
                
                if (!string.IsNullOrEmpty(text))
                {
                    if (end <= start)
                    {
                        end = start + 0.5;
                    }
                    
                    captions.Add(new Caption
                    {
                        Start = TimeSpan.FromSeconds(start),
                        End = TimeSpan.FromSeconds(end),
                        Text = text
                    });
                }
            }
        }
        
        return captions;
    }

    private static bool TryParseSrtTimestamp(string value, out TimeSpan ts)
    {
        value = value.Replace(',', '.');
        return TimeSpan.TryParseExact(
            value,
            "hh\\:mm\\:ss\\.fff",
            CultureInfo.InvariantCulture,
            out ts
        );
    }

    private static string DecodeJsonUnicode(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return value;
        }

        // Decode \uXXXX Unicode escape sequences
        return Regex.Replace(value, @"\\u([0-9A-Fa-f]{4})", match =>
        {
            int code = int.Parse(match.Groups[1].Value, NumberStyles.HexNumber);
            return char.ConvertFromUtf32(code);
        });
    }

    private static List<Caption> SplitCaptionsByWordCount(List<Caption> captions, int maxWordsPerCaption)
    {
        if (maxWordsPerCaption < 1)
        {
            maxWordsPerCaption = 1;
        }

        List<Caption> split = new List<Caption>();
        foreach (Caption caption in captions)
        {
            List<string> words = SplitWords(caption.Text);
            if (words.Count == 0)
            {
                continue;
            }

            if (words.Count <= maxWordsPerCaption)
            {
                split.Add(new Caption
                {
                    Start = caption.Start,
                    End = caption.End,
                    Text = JoinWords(words, 0, words.Count)
                });
                continue;
            }

            double startMs = caption.Start.TotalMilliseconds;
            double endMs = caption.End.TotalMilliseconds;
            if (endMs <= startMs)
            {
                endMs = startMs + 500.0;
            }

            double durationMs = endMs - startMs;
            double msPerWord = durationMs / words.Count;
            int consumed = 0;
            while (consumed < words.Count)
            {
                int chunkWordCount = maxWordsPerCaption;
                if (consumed + chunkWordCount > words.Count)
                {
                    chunkWordCount = words.Count - consumed;
                }

                double chunkStartMs = startMs + (consumed * msPerWord);
                double chunkEndMs = startMs + ((consumed + chunkWordCount) * msPerWord);
                if (consumed + chunkWordCount >= words.Count)
                {
                    chunkEndMs = endMs;
                }

                if (chunkEndMs <= chunkStartMs)
                {
                    chunkEndMs = chunkStartMs + 120.0;
                }

                split.Add(new Caption
                {
                    Start = TimeSpan.FromMilliseconds(chunkStartMs),
                    End = TimeSpan.FromMilliseconds(chunkEndMs),
                    Text = JoinWords(words, consumed, chunkWordCount)
                });

                consumed += chunkWordCount;
            }
        }

        return split;
    }

    private static List<string> SplitWords(string text)
    {
        List<string> words = new List<string>();
        if (string.IsNullOrWhiteSpace(text))
        {
            return words;
        }

        MatchCollection matches = Regex.Matches(text, @"\S+");
        foreach (Match match in matches)
        {
            string token = match.Value.Trim();
            if (token.Length > 0)
            {
                words.Add(token);
            }
        }

        return words;
    }

    private static string JoinWords(List<string> words, int startIndex, int count)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < count; i++)
        {
            if (i > 0)
            {
                sb.Append(" ");
            }

            sb.Append(words[startIndex + i]);
        }

        return sb.ToString();
    }

    private static void AddSubtitlesToTimeline(Vegas vegas, List<Caption> captions, SubtitleConfig config)
    {
        using (UndoBlock undo = new UndoBlock(vegas.Project, "Whisper Auto Subtitles"))
        {
            VideoTrack subtitleTrack = new VideoTrack(vegas.Project, vegas.Project.Tracks.Count, "Whisper Subtitles");
            if (!vegas.Project.Tracks.Contains(subtitleTrack))
            {
                vegas.Project.Tracks.Add(subtitleTrack);
            }

            PlugInNode textGenerator = FindTextGenerator(vegas);
            if (textGenerator == null)
            {
                throw new ApplicationException(
                    "Could not find a text generator plugin (Titles & Text or Legacy Text)."
                );
            }

            foreach (Caption caption in captions)
            {
                Timecode start = Timecode.FromMilliseconds(caption.Start.TotalMilliseconds);
                Timecode length = Timecode.FromMilliseconds((caption.End - caption.Start).TotalMilliseconds);

                if (length.Nanos <= 0)
                {
                    length = Timecode.FromMilliseconds(300);
                }

                Media media = new Media(textGenerator);
                MediaStream stream = media.Streams.GetItemByMediaType(MediaType.Video, 0);
                if (stream == null)
                {
                    throw new ApplicationException("Generated text media has no video stream.");
                }

                string displayText = caption.Text;
                if (config.UseLineBreaks)
                {
                    displayText = InsertLineBreaks(caption.Text, config.WordsPerLine);
                }

                TrySetGeneratedText(media, displayText, config);

                VideoEvent ev = subtitleTrack.AddVideoEvent(start, length);
                ev.Takes.Add(new Take(stream));
            }
        }
    }

    private static string InsertLineBreaks(string text, int wordsPerLine)
    {
        if (string.IsNullOrWhiteSpace(text) || wordsPerLine < 1)
        {
            return text;
        }

        string[] words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (words.Length <= wordsPerLine)
        {
            return text;
        }

        StringBuilder result = new StringBuilder();
        for (int i = 0; i < words.Length; i++)
        {
            if (i > 0)
            {
                if (i % wordsPerLine == 0)
                {
                    result.Append("\r\n");
                }
                else
                {
                    result.Append(" ");
                }
            }
            result.Append(words[i]);
        }

        return result.ToString();
    }

    private static PlugInNode FindTextGenerator(Vegas vegas)
    {
        string[] preferredNames =
        {
            "Titles & Text",
            "Legacy Text",
            "VEGAS Titles & Text",
            "Titler Pro"
        };

        foreach (string name in preferredNames)
        {
            PlugInNode found = vegas.Generators.GetChildByName(name);
            if (found != null)
            {
                return found;
            }
        }

        foreach (PlugInNode node in vegas.Generators)
        {
            if (node == null)
            {
                continue;
            }

            string n = node.Name ?? string.Empty;
            if (n.IndexOf("text", StringComparison.OrdinalIgnoreCase) >= 0 ||
                n.IndexOf("title", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return node;
            }
        }

        return null;
    }

    private static void TrySetGeneratedText(Media media, string text, SubtitleConfig config)
    {
        Effect generator = media.Generator;
        if (generator == null)
        {
            return;
        }

        // Handle OFX-based generators (Titles & Text in Vegas Pro)
        if (generator.IsOFX && generator.OFXEffect != null)
        {
            OFXEffect ofx = generator.OFXEffect;
            SetOFXText(ofx, text, config);
            return;
        }

        // Handle legacy/preset-based generators
        SetLegacyText(generator, text, config);
    }

    private static void SetOFXText(OFXEffect ofx, string text, SubtitleConfig config)
    {
        // First apply font if specified
        if (!string.IsNullOrWhiteSpace(config.FontName))
        {
            ApplyGeneratedFont(ofx, config.FontName);
        }

        // Apply outline settings
        ApplyOutlineSettings(ofx, config);

        // Vegas Titles & Text uses RTF format for the text parameter
        // We need to convert plain text to RTF
        string rtfText = ConvertToRtf(text, config);

        // Try to find the text parameter
        foreach (OFXParameter parameter in ofx.Parameters)
        {
            OFXStringParameter stringParameter = parameter as OFXStringParameter;
            if (stringParameter == null)
            {
                continue;
            }

            string name = (stringParameter.Name ?? string.Empty).ToLowerInvariant();
            string label = (stringParameter.Label ?? string.Empty).ToLowerInvariant();

            if (name.Contains("text") || label.Contains("text") || name.Contains("caption") || label.Contains("caption"))
            {
                // Check if the current value appears to be RTF
                string currentValue = stringParameter.Value ?? string.Empty;
                if (currentValue.StartsWith("{\\rtf") || currentValue.Contains("\\rtf"))
                {
                    stringParameter.Value = rtfText;
                }
                else
                {
                    // Not RTF, use plain text
                    stringParameter.Value = text;
                }
                stringParameter.ParameterChanged();
                return;
            }
        }

        // Fallback: set the first string parameter we find
        foreach (OFXParameter parameter in ofx.Parameters)
        {
            OFXStringParameter stringParameter = parameter as OFXStringParameter;
            if (stringParameter != null)
            {
                string currentValue = stringParameter.Value ?? string.Empty;
                if (currentValue.StartsWith("{\\rtf") || currentValue.Contains("\\rtf"))
                {
                    stringParameter.Value = rtfText;
                }
                else
                {
                    stringParameter.Value = text;
                }
                stringParameter.ParameterChanged();
                return;
            }
        }
    }

    private static void SetLegacyText(Effect generator, string text, SubtitleConfig config)
    {
        // For legacy text generators, try using preset with modified parameters
        // This is a fallback for older Vegas versions
        try
        {
            // Legacy generators may use a different parameter system
            // Try the Preset property if available
            if (generator.Presets != null && generator.Presets.Count > 0)
            {
                generator.Preset = generator.Presets[0].Name;
            }
        }
        catch
        {
            // Ignore errors with legacy generators
        }
    }

    private static void ApplyOutlineSettings(OFXEffect ofx, SubtitleConfig config)
    {
        foreach (OFXParameter parameter in ofx.Parameters)
        {
            string name = (parameter.Name ?? string.Empty).ToLowerInvariant();
            string label = (parameter.Label ?? string.Empty).ToLowerInvariant();

            // Try to set outline/stroke width
            if (name.Contains("outline") || name.Contains("stroke") || label.Contains("outline") || label.Contains("stroke"))
            {
                OFXDoubleParameter doubleParam = parameter as OFXDoubleParameter;
                if (doubleParam != null && (name.Contains("width") || name.Contains("size") || label.Contains("width") || label.Contains("size")))
                {
                    doubleParam.Value = config.OutlineWidth;
                    doubleParam.ParameterChanged();
                }

                // Try to set outline color via RGBA parameters
                OFXRGBAParameter rgbaParam = parameter as OFXRGBAParameter;
                if (rgbaParam != null && (name.Contains("color") || label.Contains("color")))
                {
                    rgbaParam.Value = new OFXColor(
                        config.OutlineColor.R / 255.0,
                        config.OutlineColor.G / 255.0,
                        config.OutlineColor.B / 255.0,
                        config.OutlineColor.A / 255.0
                    );
                    rgbaParam.ParameterChanged();
                }
            }
        }
    }

    private static string ConvertToRtf(string plainText, SubtitleConfig config)
    {
        if (string.IsNullOrEmpty(plainText))
        {
            plainText = " ";
        }

        // Escape RTF special characters
        string escaped = plainText
            .Replace("\\", "\\\\")
            .Replace("{", "\\{")
            .Replace("}", "\\}")
            .Replace("\n", "\\par ");

        // Build RTF string with style settings
        StringBuilder rtf = new StringBuilder();
        rtf.Append("{\\rtf1\\ansi\\deff0");
        
        // Font table
        string font = string.IsNullOrWhiteSpace(config.FontName) ? "Arial" : config.FontName.Trim();
        rtf.Append("{\\fonttbl{\\f0\\fnil\\fcharset0 ");
        rtf.Append(font);
        rtf.Append(";}}");
        
        // Color table - index 1 is text color, index 2 is outline color
        rtf.Append("{\\colortbl ;");
        rtf.AppendFormat("\\red{0}\\green{1}\\blue{2};", config.TextColor.R, config.TextColor.G, config.TextColor.B);
        rtf.AppendFormat("\\red{0}\\green{1}\\blue{2};", config.OutlineColor.R, config.OutlineColor.G, config.OutlineColor.B);
        rtf.Append("}");
        
        // Font size in half-points (24pt = 48 half-points)
        int fontSizeHalfPts = config.FontSize * 2;
        
        // Document formatting - center aligned, text color, font, size
        rtf.AppendFormat("\\pard\\qc\\cf1\\f0\\fs{0} ", fontSizeHalfPts);
        rtf.Append(escaped);
        rtf.Append("}");

        return rtf.ToString();
    }

    private static void ApplyGeneratedFont(OFXEffect ofx, string fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return;
        }

        string desired = fontName.Trim();
        string desiredLower = desired.ToLowerInvariant();

        // Try choice parameter first - this is the actual font selector in Vegas
        foreach (OFXParameter parameter in ofx.Parameters)
        {
            OFXChoiceParameter choiceParameter = parameter as OFXChoiceParameter;
            if (choiceParameter == null)
            {
                continue;
            }

            string name = (choiceParameter.Name ?? string.Empty).ToLowerInvariant();
            string label = (choiceParameter.Label ?? string.Empty).ToLowerInvariant();
            if (!(name.Contains("font") || label.Contains("font") || name.Contains("typeface") || label.Contains("typeface")))
            {
                continue;
            }

            OFXChoice[] choices = choiceParameter.Choices;
            if (choices == null)
            {
                continue;
            }

            for (int i = 0; i < choices.Length; i++)
            {
                string choiceText = choices[i] == null ? string.Empty : choices[i].ToString();
                if (choiceText.ToLowerInvariant().Contains(desiredLower))
                {
                    choiceParameter.Value = choices[i];
                    choiceParameter.ParameterChanged();
                    return;
                }
            }
        }

        // Fallback to string parameter if no choice parameter found
        foreach (OFXParameter parameter in ofx.Parameters)
        {
            OFXStringParameter stringParameter = parameter as OFXStringParameter;
            if (stringParameter == null)
            {
                continue;
            }

            string name = (stringParameter.Name ?? string.Empty).ToLowerInvariant();
            string label = (stringParameter.Label ?? string.Empty).ToLowerInvariant();
            if (name.Contains("font") || label.Contains("font") || name.Contains("typeface") || label.Contains("typeface"))
            {
                stringParameter.Value = desired;
                stringParameter.ParameterChanged();
                return;
            }
        }
    }
}
