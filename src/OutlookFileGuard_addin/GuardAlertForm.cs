using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OutlookFileGuard_addin
{
    public partial class GuardAlertForm : Form
    {
        public bool UserConfirmed { get; private set; }
        private Panel cardPanel;

        public GuardAlertForm()
        {
            InitializeForm();
            BuildUI();
            EnableFadeIn();
        }

        // =========================================================
        // FORM SETUP
        // =========================================================
        private void InitializeForm()
        {
            SuspendLayout();
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            MinimumSize = new Size(512, 100);
            BackColor = Color.FromArgb(240, 242, 245);
            ShowInTaskbar = false;
            DoubleBuffered = true;
            Padding = new Padding(18);
            Opacity = 0;
            Paint += GuardAlertForm_Paint;
            ResumeLayout();
        }

        // =========================================================
        // MAIN UI 
        // =========================================================
        private void BuildUI()
        {
            cardPanel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
            };
            Controls.Add(cardPanel);

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Color.White,
                Padding = new Padding(0)
            };

            // Header
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // Content
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // Bottoni + Footer
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            cardPanel.Controls.Add(mainLayout);

            // Aggiunta dei controlli
            mainLayout.Controls.Add(BuildHeader(), 0, 0);
            mainLayout.Controls.Add(BuildContent(), 0, 1);
            mainLayout.Controls.Add(BuildBottomPanel(), 0, 2);
        }

        // =========================================================
        // BOTTOM PANEL (Bottoni + Footer)
        // =========================================================
        private Control BuildBottomPanel()
        {
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Top,
                BackColor = Color.White,
                Padding = new Padding(28, 4, 28, 2),
                AutoSize = true,
            };

            var innerLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Color.White,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            //innerLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));  // Bottoni
            innerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // Bottoni
            innerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // Footer

            innerLayout.Controls.Add(BuildButtons(), 0, 0);
            innerLayout.Controls.Add(BuildFooter(), 0, 1);

            bottomPanel.Controls.Add(innerLayout);
            return bottomPanel;
        }

        // =========================================================
        // HEADER
        // =========================================================
        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(235, 242, 255),
                Padding = new Padding(18, 12, 18, 12)
            };

            var icon = new Label
            {
                Text = "🔒",
                AutoSize = false,
                Width = 42,
                Dock = DockStyle.Left,
                Font = new Font("Segoe UI Emoji", 20f),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var closeBtn = new Button
            {
                Text = "✕",
                Width = 34,
                Height = 34,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 10f),
                Cursor = Cursors.Hand,
                ForeColor = Color.Gray
            };
            closeBtn.FlatAppearance.BorderSize = 0;
            closeBtn.Click += (s, e) =>
            {
                UserConfirmed = false;
                Close();
            };

            var textContainer = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(10, 0, 0, 0)
            };

            var title = new Label
            {
                Text = "Quick check before sending",
                Dock = DockStyle.Top,
                AutoSize = true,
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 30, 60)
            };

            var subtitle = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top,
                AutoSize = true,
                WrapContents = true,
                Margin = new Padding(0)
            };

            var prefix = new Label
            {
                Text = "Protected by",
                Dock = DockStyle.Top,
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(80, 120, 200),
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var link = new LinkLabel
            {
                Text = "OfficeFileGuard",
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 9f),
                LinkBehavior = LinkBehavior.NeverUnderline,
                LinkColor = Color.FromArgb(50, 90, 200),
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            link.LinkClicked += (s, e) =>
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://www.officefileguard.com",
                    UseShellExecute = true
                });
            };
            subtitle.Controls.Add(prefix);
            subtitle.Controls.Add(link);

            textContainer.Controls.Add(subtitle);
            textContainer.Controls.Add(title);
            header.Controls.Add(textContainer);
            header.Controls.Add(closeBtn);
            header.Controls.Add(icon);

            return header;
        }

        // =========================================================
        // CONTENT
        // =========================================================
        private Control BuildContent()
        {
            var content = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode=AutoSizeMode.GrowAndShrink,
                Padding = new Padding(28, 10, 28, 10),
                BackColor = Color.White
            };

            var topLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode =AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1
            };
            topLayout.ColumnStyles.Clear();
            //int iconWidth = (int)(42 * DeviceDpi / 96f); // 42 è un valore “sicuro”
            //topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, iconWidth)); 
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            topLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var warnIcon = new Label
            {
                Text = "⚠",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Emoji", 30f),
                ForeColor = Color.FromArgb(230, 160, 0),
                //Height = 60,
                TextAlign = ContentAlignment.TopLeft
            };

            var messagePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoSize = true,
                Margin = new Padding(0)
            };

            messagePanel.Controls.Add(CreateTextLabel("Some attachments are marked as "));
            messagePanel.Controls.Add(CreateBoldLabel("confidential"));
            messagePanel.Controls.Add(CreateTextLabel(", "));
            messagePanel.Controls.Add(CreateTextLabel("and "));
            messagePanel.Controls.Add(CreateBoldLabel("external recipients"));
            messagePanel.Controls.Add(CreateTextLabel(" have been detected."));

            var secondMessage = new Label
            {
                Text = "To prevent accidental sharing, sending has been paused.",
                AutoSize = true,
                //MaximumSize = new Size(360, 0),
                Margin = new Padding(0, 8, 0, 0),
                Font = new Font("Segoe UI", 10f),
                ForeColor = Color.FromArgb(40, 40, 40)
            };

            var rightPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false
            };
            rightPanel.Controls.Add(messagePanel);
            rightPanel.Controls.Add(secondMessage);

            topLayout.Controls.Add(warnIcon, 0, 0);  // 1 colonna
            topLayout.Controls.Add(rightPanel, 1, 0); // 2 colonna

            // Info Box
            var infoBox = new Panel
            {
                Dock = DockStyle.Top,
                Margin = new Padding(0, 14, 0, 6),
                Padding = new Padding(16, 14, 16, 14),
                BackColor = Color.FromArgb(240, 241, 245),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var infoLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            infoLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            infoLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            infoLayout.Controls.Add(new Label
            {
                Text = "⚠ Confidential files detected",
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10f),
                ForeColor = Color.FromArgb(40, 40, 40),
                TextAlign = ContentAlignment.MiddleLeft,
            }, 0, 0);

            infoLayout.Controls.Add(new Label
            {
                Text = "✉ External recipients detected",
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10f),
                ForeColor = Color.FromArgb(40, 40, 40),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 0, 2)
            }, 0, 1);

            infoBox.Controls.Add(infoLayout);

            content.Controls.Add(infoBox);
            content.Controls.Add(topLayout);

            return content;
        }

        // =========================================================
        // BUTTONS
        // =========================================================
        private Control BuildButtons()
        {
            var container = new Panel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(0, 0, 0, 4),
                BackColor = Color.White,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            var reviewPanel = BuildButtonCard("✔ Review email", "Go back and make changes",
                Color.FromArgb(50, 90, 200), Color.White);
            reviewPanel.Margin = new Padding(0, 0, 8, 0);
            reviewPanel.Tag = new EventHandler((s, e) => { UserConfirmed = false; Close(); });

            var sendPanel = BuildButtonCard("➤ Send anyway", "I understand the risks",
                Color.FromArgb(230, 232, 236), Color.FromArgb(40, 40, 40));
            sendPanel.Margin = new Padding(8, 0, 0, 0);
            sendPanel.Tag = new EventHandler((s, e) => { UserConfirmed = true; Close(); });

            layout.Controls.Add(reviewPanel, 0, 0);
            layout.Controls.Add(sendPanel, 1, 0);

            container.Controls.Add(layout);
            return container;
        }

        // =========================================================
        // FOOTER
        // =========================================================
        private Control BuildFooter()
        {
            var footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.White,
                Padding = new Padding(0, 6, 0, 0),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            footer.Controls.Add(new Label
            {
                Text = "Like this protection? Support its development at ",
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.Gray,
                Margin = new Padding(0, 4, 0, 0)
            });

            var link = new LinkLabel
            {
                Text = "www.officefileguard.com",
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f),
                LinkColor = Color.FromArgb(50, 90, 200),
                Margin = new Padding(4, 4, 0, 0)
            };
            link.LinkClicked += (s, e) =>
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://www.officefileguard.com",
                    UseShellExecute = true
                });
            };

            footer.Controls.Add(link);
            return footer;
        }

        // =========================================================
        // BUTTON CARD (invariato)
        // =========================================================
        private Panel BuildButtonCard(string title, string subtitle, Color backColor, Color foreColor)
        {
            var panel = new Panel
            {
                Height = 56,
                Dock = DockStyle.Top,
                BackColor = backColor,
                Cursor = Cursors.Hand,
                Padding = new Padding(10)
            };

            panel.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(15, 0, 0, 0)))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1);
                }
            };

            var titleLabel = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 22,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = foreColor,
                TextAlign = ContentAlignment.MiddleCenter
            };

            var subLabel = new Label
            {
                Text = subtitle,
                Dock = DockStyle.Top,
                Height = 16,
                Font = new Font("Segoe UI", 7.5f),
                ForeColor = foreColor == Color.White ? Color.FromArgb(230, 230, 230) : Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };

            panel.Controls.Add(subLabel);
            panel.Controls.Add(titleLabel);

            foreach (Control c in panel.Controls)
            {
                c.Click += (s, e) =>
                {
                    panel.Focus();
                    if (panel.Tag is EventHandler handler)
                        handler(panel, e);
                };
            }

            return panel;
        }

        // =========================================================
        // HELPERS
        // =========================================================
        private Label CreateTextLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 10f),
                ForeColor = Color.FromArgb(40, 40, 40),
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
        }

        private Label CreateBoldLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 140, 0),
                Margin = new Padding(0)
            };
        }

        // =========================================================
        // PAINT + FADE IN + ROUNDED (invariati)
        // =========================================================
        private void GuardAlertForm_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var shadowRect = new Rectangle(8, 8, Width - 16, Height - 16);

            for (int i = 0; i < 8; i++)
            {
                using (var pen = new Pen(Color.FromArgb(12 - i, 0, 0, 0)))
                {
                    var rect = new Rectangle(shadowRect.X - i, shadowRect.Y - i,
                                           shadowRect.Width + i * 2, shadowRect.Height + i * 2);
                    using (var path = RoundedRect(rect, 18))
                        e.Graphics.DrawPath(pen, path);
                }
            }

            using (var path = RoundedRect(shadowRect, 18))
            using (var brush = new SolidBrush(Color.White))
            {
                e.Graphics.FillPath(brush, path);
                using (var borderPen = new Pen(Color.FromArgb(225, 225, 225), 1))
                    e.Graphics.DrawPath(borderPen, path);
            }

            Region = new Region(RoundedRect(shadowRect, 18));
        }

        private void EnableFadeIn()
        {
            var timer = new Timer { Interval = 15 };
            timer.Tick += (s, e) =>
            {
                if (Opacity >= 1) { timer.Stop(); timer.Dispose(); return; }
                Opacity += 0.06;
            };
            Shown += (s, e) => timer.Start();
        }

        private GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                const int CS_DROPSHADOW = 0x00020000;
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }
    }
}