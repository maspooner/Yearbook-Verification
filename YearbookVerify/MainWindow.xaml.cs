using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace YearbookVerify {
	public partial class MainWindow : System.Windows.Window, IDisposable {
		//members
		private SpellCheckCore spellCore;
		//constructors
		public MainWindow() {
			InitializeComponent();
			Icon = BitmapToBitmapSource(Properties.Resources.icon);
			spellCore = new SpellCheckCore('\n');
		}
		//methods
		/// <summary>
		/// Converts the Icon image to a BitmapSource
		/// </summary>
		private BitmapSource BitmapToBitmapSource(Bitmap b) {
			BitmapImage bitmapImage;
            using (MemoryStream memory = new MemoryStream()) {
				b.Save(memory, ImageFormat.Bmp);
				memory.Position = 0;
				bitmapImage = new BitmapImage();
				bitmapImage.BeginInit();
				bitmapImage.StreamSource = memory;
				bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
				bitmapImage.EndInit();
			}
			return bitmapImage;
		}
		public void Dispose() {
			spellCore.Dispose();
		}
		//wpf
		/// <summary>
		/// On click listener for "go" button -> check spelling and report results
		/// </summary>
		private void actionButton_Click(object sender, RoutedEventArgs e) {
			//needs some input
			if (inputBox.Text.Length > 0) {
				try {
					//check spelling of input box
					SpellingResult res = spellCore.CheckSpelling(inputBox.Text);
					//update displays
					inputBox.Text = res.UserLines;
					outputBox.Text = res.MarkedLines.Length == 0 ? "All good." : res.MarkedLines;
				}
				catch (SpellCheckException sce) {
					//spell checking failed
					MessageBox.Show(this, sce.Message, "Input in wrong format");
				}
			}
		}
		/// <summary>
		/// On click listener for reload button -> reloads all spell checking databases
		/// </summary>
		private void reloadButton_Click(object sender, RoutedEventArgs e) {
			spellCore.ReloadData();
		}
		/// <summary>
		/// Scrolls both the input and output views at the same time
		/// </summary>
		private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e) {
			ScrollViewer v = sender as ScrollViewer;
			if (sender == view1) {
				view2.ScrollToVerticalOffset(view1.VerticalOffset);
			}
			else {
				view1.ScrollToVerticalOffset(view2.VerticalOffset);
			}
		}
	}
}
