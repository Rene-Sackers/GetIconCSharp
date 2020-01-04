using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace GetIconCSharp
{
	public partial class MainWindow
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void buttonGet_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				imageIcon.Source = Icons.GetIcon(textBoxPath.Text, int.Parse(textBoxWidth.Text), int.Parse(textBoxHeight.Text));
			}
			catch (Exception)
			{
				MessageBox.Show("Could not get icon.");
			}
		}

		private void buttonSave_Click(object sender, RoutedEventArgs e)
		{
			var source = imageIcon.Source as BitmapSource;
			if (source == null) return;

			var saveFileDialog = new SaveFileDialog {AddExtension = true, Filter = "PNG Image File (*.png)|*.png"};
			if (saveFileDialog.ShowDialog() == true)
			{
				var stream = new FileStream(saveFileDialog.FileName, FileMode.CreateNew, FileAccess.Write);

				var encoder = new JpegBitmapEncoder();
				encoder.Frames.Add(BitmapFrame.Create(source));
				encoder.Save(stream);
			}
		}
	}
}