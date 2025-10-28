using System;
using System.Windows.Forms;
using System.Drawing;

namespace CorelSmartFill
{
    /// <summary>
    /// Main application window (UI)
    /// This is what the user sees and interacts with
    /// </summary>
    public partial class MainForm : Form
    {
        // ============================================
        // SECTION 1: UI CONTROLS (Buttons, Sliders, Labels, etc.)
        // ============================================

        // Connection controls
        private Button btnConnect = new Button();          // "Connect to CorelDRAW" button
        private Label lblConnectionStatus = new Label();    // Shows connection status

        // Fill controls
        private Button btnFill = new Button();             // "Fill Container" button
        private Button btnClear = new Button();            // "Clear" button
        private Button btnUndo = new Button();             // "Undo" button

        // Density control
        private Label lblDensity = new Label();            // "Density:" label
        private TrackBar sliderDensity = new TrackBar();   // Slider for density (50-5000)
        private Label lblDensityValue = new Label();       // Shows current density value

        // Spacing control
        private Label lblSpacing = new Label();            // "Spacing:" label
        private TrackBar sliderSpacing = new TrackBar();   // Slider for spacing (-100 to 200)
        private Label lblSpacingValue = new Label();       // Shows current spacing value

        // Rotation controls
        private GroupBox grpRotation = new GroupBox();     // Group box for rotation options
        private RadioButton rbNoRotation = new RadioButton();     // "No Rotation" option
        private RadioButton rbSmartRotation = new RadioButton(); // "Smart (90°)" option
        private RadioButton rbFreeRotation = new RadioButton();  // "Free (0-360°)" option

        // Overlap control
        private CheckBox chkPreventOverlap = new CheckBox(); // "Prevent Overlap" checkbox

        // Status display
        private GroupBox grpStatus = new GroupBox();       // Group box for status info
        private Label lblStatusCount = new Label();        // Shows element count
        private Label lblStatusCoverage = new Label();     // Shows coverage percentage
        private ProgressBar progressBar = new ProgressBar(); // Shows operation progress

        // ============================================
        // SECTION 2: BACKEND OBJECTS
        // ============================================

        // CorelDRAW connection object
        private CorelDRAWConnection corelConnection = new CorelDRAWConnection();

        // Packing algorithm object
        private PackingAlgorithm? packingAlgorithm = null;

        // Stored references for operations
        private dynamic? currentContainer = null;
        private System.Collections.Generic.List<dynamic>? currentFillElements = null;

        /// <summary>
        /// Constructor - called when form is created
        /// Sets up all UI controls and event handlers
        /// </summary>
        public MainForm()
        {
            // ============================================
            // STEP 1: Basic form setup
            // ============================================

            // Set window title
            this.Text = "CorelDRAW Smart Fill Tool v1.0";

            // Set window size (width x height in pixels)
            this.ClientSize = new Size(450, 650);

            // Prevent resizing (keeps layout clean)
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // Center window on screen when it opens
            this.StartPosition = FormStartPosition.CenterScreen;

            // ============================================
            // STEP 2: Set up all UI controls
            // ============================================

            SetupConnectionControls();  // Top section - connection button
            SetupFillControls();        // Fill operation buttons
            SetupDensityControls();     // Density slider
            SetupSpacingControls();     // Spacing slider
            SetupRotationControls();    // Rotation options
            SetupOverlapControls();     // Overlap checkbox
            SetupStatusControls();      // Bottom status display

            // ============================================
            // STEP 3: Initial state
            // ============================================

            // Disable fill controls until connected
            EnableFillControls(false);
        }

        /// <summary>
        /// Set up connection controls (top of form)
        /// </summary>
        private void SetupConnectionControls()
        {
            // CONNECT BUTTON
            btnConnect.Text = "Connect to CorelDRAW";
            btnConnect.Location = new Point(20, 20);      // X=20, Y=20 pixels from top-left
            btnConnect.Size = new Size(200, 35);          // Width=200, Height=35 pixels
            btnConnect.Click += BtnConnect_Click;         // When clicked, call BtnConnect_Click method
            this.Controls.Add(btnConnect);                // Add button to form

            // CONNECTION STATUS LABEL
            lblConnectionStatus.Text = "Not Connected";
            lblConnectionStatus.Location = new Point(230, 25);
            lblConnectionStatus.Size = new Size(200, 25);
            lblConnectionStatus.ForeColor = Color.Red;    // Red text for "not connected"
            this.Controls.Add(lblConnectionStatus);
        }

        /// <summary>
        /// Set up fill operation buttons
        /// </summary>
        private void SetupFillControls()
        {
            // FILL BUTTON (main action button)
            btnFill.Text = "Fill Container";
            btnFill.Location = new Point(20, 70);
            btnFill.Size = new Size(130, 40);
            btnFill.BackColor = Color.FromArgb(0, 120, 215); // Blue button
            btnFill.ForeColor = Color.White;                  // White text
            btnFill.FlatStyle = FlatStyle.Flat;               // Modern flat style
            btnFill.Click += BtnFill_Click;
            this.Controls.Add(btnFill);

            // CLEAR BUTTON
            btnClear.Text = "Clear";
            btnClear.Location = new Point(160, 70);
            btnClear.Size = new Size(130, 40);
            btnClear.Click += BtnClear_Click;
            this.Controls.Add(btnClear);

            // UNDO BUTTON
            btnUndo.Text = "Undo";
            btnUndo.Location = new Point(300, 70);
            btnUndo.Size = new Size(130, 40);
            btnUndo.Click += BtnUndo_Click;
            this.Controls.Add(btnUndo);
        }

        /// <summary>
        /// Set up density slider control
        /// </summary>
        private void SetupDensityControls()
        {
            // LABEL
            lblDensity.Text = "Fill Density (attempts):";
            lblDensity.Location = new Point(20, 130);
            lblDensity.Size = new Size(200, 20);
            this.Controls.Add(lblDensity);

            // SLIDER (TrackBar)
            sliderDensity.Location = new Point(20, 155);
            sliderDensity.Size = new Size(350, 45);
            sliderDensity.Minimum = 50;           // Minimum value: 50 attempts
            sliderDensity.Maximum = 5000;         // Maximum value: 5000 attempts
            sliderDensity.Value = 500;            // Default: 500 attempts
            sliderDensity.TickFrequency = 500;    // Show tick marks every 500
            sliderDensity.ValueChanged += SliderDensity_ValueChanged;  // Update label when changed
            this.Controls.Add(sliderDensity);

            // VALUE LABEL (shows current slider value)
            lblDensityValue.Text = "500";
            lblDensityValue.Location = new Point(380, 155);
            lblDensityValue.Size = new Size(50, 20);
            lblDensityValue.TextAlign = ContentAlignment.MiddleRight;
            this.Controls.Add(lblDensityValue);
        }

        /// <summary>
        /// Set up spacing slider control
        /// </summary>
        private void SetupSpacingControls()
        {
            // LABEL
            lblSpacing.Text = "Minimum Spacing (%):";
            lblSpacing.Location = new Point(20, 210);
            lblSpacing.Size = new Size(200, 20);
            this.Controls.Add(lblSpacing);

            // SLIDER
            sliderSpacing.Location = new Point(20, 235);
            sliderSpacing.Size = new Size(350, 45);
            sliderSpacing.Minimum = -100;         // -100% = lots of overlap
            sliderSpacing.Maximum = 200;          // 200% = very sparse
            sliderSpacing.Value = 0;              // Default: 0% (touching)
            sliderSpacing.TickFrequency = 25;     // Tick marks every 25%
            sliderSpacing.ValueChanged += SliderSpacing_ValueChanged;
            this.Controls.Add(sliderSpacing);

            // VALUE LABEL
            lblSpacingValue.Text = "0%";
            lblSpacingValue.Location = new Point(380, 235);
            lblSpacingValue.Size = new Size(50, 20);
            lblSpacingValue.TextAlign = ContentAlignment.MiddleRight;
            this.Controls.Add(lblSpacingValue);
        }

        /// <summary>
        /// Set up rotation options (radio buttons)
        /// </summary>
        private void SetupRotationControls()
        {
            // GROUP BOX (container for radio buttons)
            grpRotation.Text = "Rotation Mode";
            grpRotation.Location = new Point(20, 290);
            grpRotation.Size = new Size(410, 100);
            this.Controls.Add(grpRotation);

            // RADIO BUTTON 1: No Rotation
            rbNoRotation.Text = "No Rotation - Keep original angle";
            rbNoRotation.Location = new Point(10, 25);
            rbNoRotation.Size = new Size(380, 20);
            grpRotation.Controls.Add(rbNoRotation);

            // RADIO BUTTON 2: Smart Rotation (default)
            rbSmartRotation.Text = "Smart Rotation - 90° increments (recommended)";
            rbSmartRotation.Location = new Point(10, 50);
            rbSmartRotation.Size = new Size(380, 20);
            rbSmartRotation.Checked = true;  // This is selected by default
            grpRotation.Controls.Add(rbSmartRotation);

            // RADIO BUTTON 3: Free Rotation
            rbFreeRotation.Text = "Free Rotation - Any angle 0-360°";
            rbFreeRotation.Location = new Point(10, 75);
            rbFreeRotation.Size = new Size(380, 20);
            grpRotation.Controls.Add(rbFreeRotation);
        }

        /// <summary>
        /// Set up overlap prevention checkbox
        /// </summary>
        private void SetupOverlapControls()
        {
            chkPreventOverlap.Text = "Prevent Overlap (slower, but cleaner result)";
            chkPreventOverlap.Location = new Point(20, 405);
            chkPreventOverlap.Size = new Size(400, 20);
            chkPreventOverlap.Checked = true;  // Enabled by default
            this.Controls.Add(chkPreventOverlap);
        }

        /// <summary>
        /// Set up status display (bottom of form)
        /// </summary>
        private void SetupStatusControls()
        {
            // GROUP BOX
            grpStatus.Text = "Status";
            grpStatus.Location = new Point(20, 440);
            grpStatus.Size = new Size(410, 120);
            this.Controls.Add(grpStatus);

            // ELEMENT COUNT LABEL
            lblStatusCount.Text = "Elements placed: 0";
            lblStatusCount.Location = new Point(10, 25);
            lblStatusCount.Size = new Size(380, 20);
            grpStatus.Controls.Add(lblStatusCount);

            // COVERAGE LABEL
            lblStatusCoverage.Text = "Coverage: 0%";
            lblStatusCoverage.Location = new Point(10, 50);
            lblStatusCoverage.Size = new Size(380, 20);
            grpStatus.Controls.Add(lblStatusCoverage);

            // PROGRESS BAR
            progressBar.Location = new Point(10, 80);
            progressBar.Size = new Size(390, 25);
            progressBar.Style = ProgressBarStyle.Continuous;
            grpStatus.Controls.Add(progressBar);
        }

        // ============================================
        // SECTION 3: EVENT HANDLERS (what happens when user clicks buttons)
        // ============================================

        /// <summary>
        /// Called when "Connect to CorelDRAW" button is clicked
        /// </summary>
        private void BtnConnect_Click(object? sender, EventArgs e)
        {
            try
            {
                // Show "connecting..." message
                btnConnect.Enabled = false;
                btnConnect.Text = "Connecting...";
                Application.DoEvents(); // Update UI immediately

                // Attempt connection
                if (corelConnection.Connect())
                {
                    // SUCCESS!
                    lblConnectionStatus.Text = "Connected";
                    lblConnectionStatus.ForeColor = Color.Green;
                    btnConnect.Text = "Connected ✓";

                    // Create packing algorithm object
                    packingAlgorithm = new PackingAlgorithm(corelConnection);

                    // Enable fill controls
                    EnableFillControls(true);

                    MessageBox.Show(
                        "Successfully connected to CorelDRAW!\n\n" +
                        "Now:\n" +
                        "1. Select your fill element(s) and container in CorelDRAW\n" +
                        "2. Adjust settings in this window\n" +
                        "3. Click 'Fill Container'",
                        "Connected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // CONNECTION FAILED
                lblConnectionStatus.Text = "Connection Failed";
                lblConnectionStatus.ForeColor = Color.Red;
                btnConnect.Text = "Connect to CorelDRAW";
                btnConnect.Enabled = true;

                MessageBox.Show(
                    $"Failed to connect to CorelDRAW:\n\n{ex.Message}\n\n" +
                    "Make sure:\n" +
                    "1. CorelDRAW is running\n" +
                    "2. A document is open\n" +
                    "3. You have CorelDRAW 2018 or later",
                    "Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Called when "Fill Container" button is clicked
        /// This is the main action - it performs the fill operation
        /// </summary>
        private void BtnFill_Click(object? sender, EventArgs e)
        {
            if (packingAlgorithm == null)
            {
                MessageBox.Show("Not connected to CorelDRAW", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // STEP 1: Get selected shapes from CorelDRAW
                btnFill.Enabled = false;
                btnFill.Text = "Getting selection...";
                Application.DoEvents();

                var (container, fillElements) = corelConnection.GetSelectedShapes();
                currentContainer = container;
                currentFillElements = fillElements;

                // STEP 2: Get settings from UI
                int density = sliderDensity.Value;
                double spacing = sliderSpacing.Value / 100.0;  // Convert percentage to decimal
                bool allowRotation = rbSmartRotation.Checked || rbFreeRotation.Checked;
                bool smartRotation = rbSmartRotation.Checked;
                bool preventOverlap = chkPreventOverlap.Checked;

                // STEP 3: Show progress
                btnFill.Text = "Filling...";
                progressBar.Style = ProgressBarStyle.Marquee; // Animated progress bar
                Application.DoEvents();

                // STEP 4: Perform the fill operation
                int placedCount = packingAlgorithm.FillContainer(
                    container,
                    fillElements,
                    density,
                    spacing,
                    allowRotation,
                    smartRotation,
                    preventOverlap);

                // STEP 5: Update status display
                var stats = packingAlgorithm.GetStatistics(container);
                lblStatusCount.Text = $"Elements placed: {stats.count}";
                lblStatusCoverage.Text = $"Coverage: {stats.coverage:F1}%";
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = Math.Min(100, (int)stats.coverage);

                // STEP 6: Show success message
                MessageBox.Show(
                    $"Fill operation complete!\n\n" +
                    $"Elements placed: {placedCount}\n" +
                    $"Coverage: {stats.coverage:F1}%\n\n" +
                    $"You can now adjust density using the slider,\n" +
                    $"or click Undo to remove all elements.",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                btnFill.Text = "Fill Container";
                btnFill.Enabled = true;
            }
            catch (Exception ex)
            {
                // ERROR during fill operation
                MessageBox.Show(
                    $"Fill operation failed:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                btnFill.Text = "Fill Container";
                btnFill.Enabled = true;
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
            }
        }

        /// <summary>
        /// Called when density slider value changes
        /// Updates the label to show current value
        /// </summary>
        private void SliderDensity_ValueChanged(object? sender, EventArgs e)
        {
            lblDensityValue.Text = sliderDensity.Value.ToString();
        }

        /// <summary>
        /// Called when spacing slider value changes
        /// Updates the label to show current value
        /// </summary>
        private void SliderSpacing_ValueChanged(object? sender, EventArgs e)
        {
            lblSpacingValue.Text = $"{sliderSpacing.Value}%";
        }

        /// <summary>
        /// Called when "Clear" button is clicked
        /// Clears internal state (doesn't delete shapes from CorelDRAW)
        /// </summary>
        private void BtnClear_Click(object? sender, EventArgs e)
        {
            if (packingAlgorithm != null)
            {
                packingAlgorithm.Clear();
                lblStatusCount.Text = "Elements placed: 0";
                lblStatusCoverage.Text = "Coverage: 0%";
                progressBar.Value = 0;

                MessageBox.Show("Internal state cleared.\n\n" +
                    "Note: This doesn't delete shapes from CorelDRAW.\n" +
                    "Use Undo (Ctrl+Z) in CorelDRAW to remove shapes.",
                    "Cleared",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Called when "Undo" button is clicked
        /// This would trigger undo in CorelDRAW
        /// </summary>
        private void BtnUndo_Click(object? sender, EventArgs e)
        {
            MessageBox.Show(
                "To undo the fill operation:\n\n" +
                "Press Ctrl+Z in CorelDRAW\n\n" +
                "All placed elements will be removed in one step.",
                "Undo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// Enable or disable fill controls
        /// Called after connection succeeds/fails
        /// </summary>
        private void EnableFillControls(bool enabled)
        {
            btnFill.Enabled = enabled;
            btnClear.Enabled = enabled;
            btnUndo.Enabled = enabled;
            sliderDensity.Enabled = enabled;
            sliderSpacing.Enabled = enabled;
            rbNoRotation.Enabled = enabled;
            rbSmartRotation.Enabled = enabled;
            rbFreeRotation.Enabled = enabled;
            chkPreventOverlap.Enabled = enabled;
        }

        /// <summary>
        /// Called when form is closing
        /// Clean up resources and disconnect from CorelDRAW
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            // Disconnect from CorelDRAW (releases COM objects)
            corelConnection.Disconnect();
        }
    }
}
