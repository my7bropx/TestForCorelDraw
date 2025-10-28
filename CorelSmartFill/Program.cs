using System;
using System.Windows.Forms;

namespace CorelSmartFill
{
    /// <summary>
    /// Main program entry point
    /// This is where the application starts when you run the .exe file
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// [STAThread] means "Single Threaded Apartment" - required for COM interop with CorelDRAW
        /// Without this, COM automation will fail
        /// </summary>
        [STAThread]
        static void Main()
        {
            // STEP 1: Enable visual styles for modern Windows look
            // This makes buttons, textboxes, etc. look modern (Windows 10/11 style)
            Application.EnableVisualStyles();

            // STEP 2: Use compatible text rendering
            // This ensures text in the UI looks crisp and clear
            Application.SetCompatibleTextRenderingDefault(false);

            // STEP 3: Set up global exception handler
            // If any unhandled error occurs, we'll catch it and show a friendly message
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;

            // STEP 4: Create and run the main form (the UI window)
            // This shows the window and keeps it running until user closes it
            Application.Run(new MainForm());
        }

        /// <summary>
        /// Global exception handler
        /// This catches any errors that weren't handled elsewhere
        /// Instead of crashing, we show a user-friendly error message
        /// </summary>
        /// <param name="sender">The object that threw the exception</param>
        /// <param name="e">Details about the exception</param>
        private static void Application_ThreadException(object sender,
            System.Threading.ThreadExceptionEventArgs e)
        {
            // Show error message to user
            MessageBox.Show(
                $"An error occurred:\n\n{e.Exception.Message}\n\nStack Trace:\n{e.Exception.StackTrace}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
