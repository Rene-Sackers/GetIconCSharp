using System;
using System.ComponentModel;
using System.Windows;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;

namespace GetIconCSharp
{
	internal static class Icons
	{
		[Flags]
		public enum Siigbf
		{
			SiigbfResizeToFit = 0,
			SiigbfBiggerSizeOk = 1,
			SiigbfMemoryOnly = 2,
			SiigbfIconOnly = 4,
			SiigbfThumbnailOnly = 8,
			SiigbfIncacheOnly = 16
		}

		public enum Sigdn : uint
		{
			NormalDisplay = 0,
			ParentRelativeParsing = 0x80018001u,
			ParentRelativeForAddressbar = 0x8001c001u,
			DesktopAbsoluteParsing = 0x80028000u,
			ParentRelativeEditing = 0x80031001u,
			DesktopAbsoluteEditing = 0x8004c000u,
			FilesysPath = 0x80058000u,
			Url = 0x80068000u
		}

		[ComImportAttribute]
		[GuidAttribute("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
		[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
		public interface IShellItemImageFactory
		{
			void GetImage(SizeStruct sizeStruct, Siigbf flags, ref IntPtr phbm);
		}

		[ComImport]
		[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		public interface IShellItem
		{
			void BindToHandler(
				IntPtr pbc,
				[MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
				ref IntPtr ppv);

			void GetParent(ref IShellItem ppsi);
			void GetDisplayName(Sigdn sigdnName, ref IntPtr ppszName);
			void GetAttributes(uint sfgaoMask, ref uint psfgaoAttribs);
			void Compare(IShellItem psi, uint hint, ref int piOrder);
		}

		[DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
		public static extern void SHCreateItemFromParsingName(
			[MarshalAs(UnmanagedType.LPWStr)] string pszPath,
			IntPtr pbc,
			[MarshalAs(UnmanagedType.LPStruct)] Guid riid,
			[MarshalAs(UnmanagedType.Interface, IidParameterIndex = 2)]
			ref IShellItem ppv);

		[DllImport("gdi32.dll", SetLastError = true)]
		private static extern bool DeleteObject(IntPtr hObject);

		[StructLayout(LayoutKind.Sequential)]
		public struct SizeStruct
		{
			public int cx;
			public int cy;

			public SizeStruct(int cx, int cy)
			{
				this.cx = cx;
				this.cy = cy;
			}
		}

		public static BitmapSource GetIcon(string path) => GetIcon(path, 256, 256);

		public static BitmapSource GetIcon(string path, int width, int height)
		{
			IShellItem ppsi = null;
			var hBitmap = IntPtr.Zero;

			SHCreateItemFromParsingName(
				path,
				IntPtr.Zero,
				new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"),
				ref ppsi
			);

			((IShellItemImageFactory) ppsi).GetImage(
				new SizeStruct(width, height),
				Siigbf.SiigbfResizeToFit,
				ref hBitmap
			);

			var bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
				hBitmap,
				IntPtr.Zero,
				Int32Rect.Empty,
				BitmapSizeOptions.FromEmptyOptions()
			);

			if (!DeleteObject(hBitmap)) throw new Win32Exception();
			Marshal.ReleaseComObject(ppsi);

			return bitmapSource;
		}
	}
}