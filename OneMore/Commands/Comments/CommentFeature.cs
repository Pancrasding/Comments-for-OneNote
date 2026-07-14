// SPDX-License-Identifier: MPL-2.0
// Copyright 2026 Comments for OneNote contributors
// Embedded comments for OneNote desktop

namespace River.OneMoreAddIn.Commands
{
	using Newtonsoft.Json;
	using River.OneMoreAddIn.Models;
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.Linq;
	using System.Runtime.InteropServices;
	using System.Threading;
	using System.Threading.Tasks;
	using System.Windows.Forms;
	using System.Xml.Linq;

	internal sealed class CommentRecord
	{
		public string Id { get; set; }
		public string PageId { get; set; }
		public string AnchorObjectId { get; set; }
		public string Quote { get; set; }
		public string Prefix { get; set; }
		public string Suffix { get; set; }
		public string Comment { get; set; }
		public bool Resolved { get; set; }
		public DateTime CreatedUtc { get; set; }
		public DateTime UpdatedUtc { get; set; }

		[JsonIgnore]
		public bool Attached { get; set; }
	}


	internal static class CommentStore
	{
		public const string MetaName = "OneMoreCommentsV1";

		public static List<CommentRecord> Load(Page page)
		{
			var json = page?.GetMetaContent(MetaName);
			if (string.IsNullOrWhiteSpace(json))
			{
				return new List<CommentRecord>();
			}

			try
			{
				return JsonConvert.DeserializeObject<List<CommentRecord>>(json)
					?? new List<CommentRecord>();
			}
			catch (Exception exc)
			{
				Logger.Current.WriteLine("cannot read embedded comments metadata", exc);
				return new List<CommentRecord>();
			}
		}

		public static void Save(Page page, IEnumerable<CommentRecord> comments)
		{
			page.SetMeta(MetaName, JsonConvert.SerializeObject(
				comments, Formatting.None,
				new JsonSerializerSettings { DateTimeZoneHandling = DateTimeZoneHandling.Utc }));
		}

		public static string PlainText(XElement paragraph)
		{
			if (paragraph is null)
			{
				return string.Empty;
			}

			var ns = paragraph.GetDefaultNamespace();
			if (ns == XNamespace.None)
			{
				ns = paragraph.GetNamespaceOfPrefix(OneNote.Prefix);
			}

			return string.Concat(paragraph.Descendants(ns + "T").Select(t =>
			{
				try { return t.GetCData().GetWrapper().Value; }
				catch { return t.Value; }
			}));
		}

		public static XElement FindParagraph(Page page, string objectId)
		{
			if (page is null || string.IsNullOrEmpty(objectId))
			{
				return null;
			}

			return page.Root.Descendants(page.Namespace + "OE")
				.FirstOrDefault(e => e.Attribute("objectID")?.Value == objectId);
		}

		public static bool Rebind(Page page, CommentRecord record)
		{
			var current = FindParagraph(page, record.AnchorObjectId);
			if (current is not null && PlainText(current).Contains(record.Quote))
			{
				record.Attached = true;
				return false;
			}

			var candidates = page.Root.Descendants(page.Namespace + "OE")
				.Where(e => e.Attribute("objectID") is not null)
				.Select(e => new { Element = e, Text = PlainText(e) })
				.Where(x => x.Text.Contains(record.Quote))
				.Select(x => new
				{
					x.Element,
					Score = Score(x.Text, record)
				})
				.OrderByDescending(x => x.Score)
				.ToList();

			if (candidates.Count == 0)
			{
				record.Attached = false;
				return false;
			}

			var best = candidates[0];
			if (candidates.Count > 1 && best.Score == candidates[1].Score && best.Score == 0)
			{
				record.Attached = false;
				return false;
			}

			var newId = best.Element.Attribute("objectID")?.Value;
			var changed = newId != record.AnchorObjectId;
			record.AnchorObjectId = newId;
			record.Attached = true;
			if (changed)
			{
				record.UpdatedUtc = DateTime.UtcNow;
			}

			return changed;
		}

		private static int Score(string text, CommentRecord record)
		{
			var score = 0;
			if (!string.IsNullOrEmpty(record.Prefix) && text.Contains(record.Prefix + record.Quote))
			{
				score += 4;
			}
			if (!string.IsNullOrEmpty(record.Suffix) && text.Contains(record.Quote + record.Suffix))
			{
				score += 4;
			}
			if (!string.IsNullOrEmpty(record.Prefix) && text.Contains(record.Prefix))
			{
				score++;
			}
			if (!string.IsNullOrEmpty(record.Suffix) && text.Contains(record.Suffix))
			{
				score++;
			}
			return score;
		}
	}


	internal sealed class AddCommentCommand : Command
	{
		public override async Task Execute(params object[] args)
		{
			await using var one = new OneNote(out var page, out _);
			if (page is null || !page.IsValid)
			{
				ShowError("无法读取当前 OneNote 页面。");
				return;
			}

			var range = new Models.SelectionRange(page);
			var selected = range.GetSelections().ToList();
			if (selected.Count == 0 || (range.Scope == SelectionScope.None || range.Scope == SelectionScope.TextCursor))
			{
				ShowInfo("请先选中要评论的一段文字。");
				return;
			}

			if (!range.SingleParagraph)
			{
				ShowInfo("当前版本请在同一个段落内选择文字；跨段评论稍后再支持。");
				return;
			}

			var quote = new PageEditor(page).GetSelectedText()?.Trim();
			if (string.IsNullOrWhiteSpace(quote))
			{
				ShowInfo("选中的内容不是可评论的文字。");
				return;
			}

			using var dialog = new CommentEditDialog(quote, string.Empty);
			if (dialog.ShowDialog(owner) != DialogResult.OK)
			{
				IsCancelled = true;
				return;
			}

			var paragraph = selected[0].Ancestors(page.Namespace + "OE").FirstOrDefault();
			var paragraphText = CommentStore.PlainText(paragraph);
			var index = paragraphText.IndexOf(quote, StringComparison.Ordinal);
			var prefix = index > 0
				? paragraphText.Substring(Math.Max(0, index - 40), Math.Min(40, index))
				: string.Empty;
			var suffixStart = index < 0 ? -1 : index + quote.Length;
			var suffix = suffixStart >= 0 && suffixStart < paragraphText.Length
				? paragraphText.Substring(suffixStart, Math.Min(40, paragraphText.Length - suffixStart))
				: string.Empty;

			var comments = CommentStore.Load(page);
			var now = DateTime.UtcNow;
			comments.Add(new CommentRecord
			{
				Id = Guid.NewGuid().ToString("N"),
				PageId = page.PageId,
				AnchorObjectId = paragraph?.Attribute("objectID")?.Value,
				Quote = quote,
				Prefix = prefix,
				Suffix = suffix,
				Comment = dialog.CommentText,
				Resolved = false,
				CreatedUtc = now,
				UpdatedUtc = now,
				Attached = true
			});

			new PageEditor(page).EditSelected(node =>
			{
				if (node is XText text)
				{
					return new XElement("span",
						new XAttribute("style", "background:#FFF2CC"), text);
				}

				var span = (XElement)node;
				span.GetAttributeValue("style", out var style, string.Empty);
				if (!style.Contains("background:"))
				{
					span.SetAttributeValue("style", $"{style};background:#FFF2CC".TrimStart(';'));
				}
				return span;
			});

			CommentStore.Save(page, comments);
			await one.Update(page);
			CommentPaneManager.Show();
			CommentPaneManager.Refresh();
		}
	}


	internal sealed class CommentPaneCommand : Command
	{
		public override async Task Execute(params object[] args)
		{
			CommentPaneManager.Toggle();
			await Task.Yield();
		}
	}


	internal sealed class CommentEditDialog : Form
	{
		private readonly TextBox editor;

		public CommentEditDialog(string quote, string value)
		{
			Text = "OneNote 评论";
			StartPosition = FormStartPosition.CenterParent;
			FormBorderStyle = FormBorderStyle.FixedDialog;
			MinimizeBox = false;
			MaximizeBox = false;
			ShowInTaskbar = false;
			Width = 520;
			Height = 330;
			Font = new Font("Microsoft YaHei UI", 9F);

			var quoteLabel = new Label
			{
				Text = "原文：" + quote,
				AutoEllipsis = true,
				Dock = DockStyle.Top,
				Height = 48,
				Padding = new Padding(12, 12, 12, 4)
			};
			editor = new TextBox
			{
				Text = value ?? string.Empty,
				Multiline = true,
				AcceptsReturn = true,
				ScrollBars = ScrollBars.Vertical,
				Dock = DockStyle.Fill,
				Margin = new Padding(12)
			};

			var buttons = new FlowLayoutPanel
			{
				Dock = DockStyle.Bottom,
				Height = 50,
				FlowDirection = FlowDirection.RightToLeft,
				Padding = new Padding(8)
			};
			var save = new Button { Text = "保存", DialogResult = DialogResult.OK, Width = 86 };
			var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 86 };
			buttons.Controls.Add(save);
			buttons.Controls.Add(cancel);

			var center = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 4, 12, 8) };
			center.Controls.Add(editor);
			Controls.Add(center);
			Controls.Add(buttons);
			Controls.Add(quoteLabel);
			AcceptButton = save;
			CancelButton = cancel;
		}

		public string CommentText => editor.Text.Trim();

		protected override void OnFormClosing(FormClosingEventArgs e)
		{
			if (DialogResult == DialogResult.OK && string.IsNullOrWhiteSpace(CommentText))
			{
				MessageBox.Show(this, "评论内容不能为空。", "OneNote 评论",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				e.Cancel = true;
				return;
			}
			base.OnFormClosing(e);
		}
	}


	internal static class CommentPaneManager
	{
		private static readonly object gate = new();
		private static CommentPaneForm pane;
		private static Thread thread;

		public static void Toggle()
		{
			lock (gate)
			{
				if (pane is not null && !pane.IsDisposed)
				{
					pane.BeginInvoke(new Action(pane.Close));
					return;
				}
			}
			Show();
		}

		public static void Show()
		{
			lock (gate)
			{
				if (pane is not null && !pane.IsDisposed)
				{
					pane.BeginInvoke(new Action(() => pane.ShowPane()));
					return;
				}

				thread = new Thread(() =>
				{
					var form = new CommentPaneForm();
					lock (gate) { pane = form; }
					form.FormClosed += (_, _) =>
					{
						lock (gate) { pane = null; thread = null; }
					};
					Application.Run(form);
				})
				{
					Name = "OneMoreCommentsPane",
					IsBackground = true
				};
				thread.SetApartmentState(ApartmentState.STA);
				thread.Start();
			}
		}

		public static void Refresh()
		{
			lock (gate)
			{
				if (pane is not null && !pane.IsDisposed && pane.IsHandleCreated)
				{
					pane.BeginInvoke(new Action(() => _ = pane.RefreshComments()));
				}
			}
		}

		public static void Close()
		{
			lock (gate)
			{
				if (pane is not null && !pane.IsDisposed && pane.IsHandleCreated)
				{
					pane.BeginInvoke(new Action(pane.Close));
				}
			}
		}
	}


	internal sealed class CommentPaneForm : Form
	{
		private const int PaneWidth = 390;
		private readonly FlowLayoutPanel cards;
		private readonly Label pageLabel;
		private readonly System.Windows.Forms.Timer timer;
		private readonly SemaphoreSlim refreshing = new(1, 1);
		private IntPtr host;
		private string pageId;

		public CommentPaneForm()
		{
			Text = "OneNote 评论";
			FormBorderStyle = FormBorderStyle.None;
			ShowInTaskbar = false;
			TopMost = false;
			BackColor = Color.FromArgb(248, 249, 251);
			Font = new Font("Microsoft YaHei UI", 9F);
			Width = PaneWidth;

			var header = new Panel
			{
				Dock = DockStyle.Top,
				Height = 86,
				BackColor = Color.White,
				Padding = new Padding(12, 8, 8, 8)
			};
			var title = new Label
			{
				Text = "评论",
				Font = new Font(Font.FontFamily, 13F, FontStyle.Bold),
				AutoSize = true,
				Location = new Point(12, 10)
			};
			pageLabel = new Label
			{
				Text = "正在读取当前页面…",
				AutoEllipsis = true,
				Location = new Point(13, 45),
				Width = 245,
				Height = 24,
				ForeColor = Color.DimGray
			};
			var add = new Button { Text = "+ 添加", Width = 68, Height = 28, Location = new Point(268, 9) };
			var close = new Button { Text = "×", Width = 34, Height = 28, Location = new Point(342, 9) };
			var refresh = new Button { Text = "刷新", Width = 68, Height = 26, Location = new Point(306, 47) };
			add.Click += async (_, _) =>
			{
				if (AddIn.Self is not null)
				{
					await AddIn.Self.AddCommentCmd(null);
				}
			};
			close.Click += (_, _) => Close();
			refresh.Click += async (_, _) => await RefreshComments();
			header.Controls.Add(title);
			header.Controls.Add(pageLabel);
			header.Controls.Add(add);
			header.Controls.Add(close);
			header.Controls.Add(refresh);

			cards = new FlowLayoutPanel
			{
				Dock = DockStyle.Fill,
				FlowDirection = FlowDirection.TopDown,
				WrapContents = false,
				AutoScroll = true,
				Padding = new Padding(10)
			};
			Controls.Add(cards);
			Controls.Add(header);

			timer = new System.Windows.Forms.Timer { Interval = 2000 };
			timer.Tick += async (_, _) =>
			{
				AttachAndPosition();
				await RefreshComments();
			};
		}

		protected override bool ShowWithoutActivation => false;

		protected override async void OnShown(EventArgs e)
		{
			base.OnShown(e);
			AttachAndPosition();
			timer.Start();
			await RefreshComments();
		}

		protected override void OnFormClosed(FormClosedEventArgs e)
		{
			timer.Stop();
			refreshing.Dispose();
			base.OnFormClosed(e);
		}

		public void ShowPane()
		{
			AttachAndPosition();
			Show();
			BringToFront();
			_ = RefreshComments();
		}

		private void AttachAndPosition()
		{
			try
			{
				using var one = new OneNote();
				var candidate = one.WindowHandle;
				if (candidate == IntPtr.Zero)
				{
					return;
				}

				if (host != candidate || PaneNative.GetParent(Handle) != candidate)
				{
					host = candidate;
					PaneNative.SetParent(Handle, host);
					var style = PaneNative.GetWindowLong(Handle, PaneNative.GWL_STYLE);
					style &= ~(PaneNative.WS_POPUP | PaneNative.WS_CAPTION | PaneNative.WS_THICKFRAME);
					style |= PaneNative.WS_CHILD | PaneNative.WS_VISIBLE | PaneNative.WS_CLIPSIBLINGS;
					PaneNative.SetWindowLong(Handle, PaneNative.GWL_STYLE, style);
				}

				if (PaneNative.GetClientRect(host, out var rect))
				{
					var width = Math.Min(PaneWidth, Math.Max(300, (rect.Right - rect.Left) / 3));
					var height = Math.Max(200, rect.Bottom - rect.Top);
					PaneNative.SetWindowPos(Handle, IntPtr.Zero,
						Math.Max(0, rect.Right - width), 0, width, height,
						PaneNative.SWP_NOZORDER | PaneNative.SWP_NOACTIVATE | PaneNative.SWP_SHOWWINDOW);
				}
			}
			catch (Exception exc)
			{
				Logger.Current.WriteLine("cannot attach comments pane", exc);
			}
		}

		public async Task RefreshComments(bool onlyIfPageChanged = false)
		{
			if (!await refreshing.WaitAsync(0))
			{
				return;
			}

			try
			{
				await using var one = new OneNote();
				var current = one.CurrentPageId;
				if (string.IsNullOrEmpty(current))
				{
					pageLabel.Text = "没有活动页面";
					return;
				}
				if (onlyIfPageChanged && current == pageId)
				{
					return;
				}

				pageId = current;
				var page = await one.GetPage(current, OneNote.PageDetail.All);
				if (page is null)
				{
					return;
				}

				var comments = CommentStore.Load(page);
				var changed = false;
				foreach (var comment in comments)
				{
					changed |= CommentStore.Rebind(page, comment);
				}
				if (changed)
				{
					CommentStore.Save(page, comments);
					await one.Update(page);
				}

				pageLabel.Text = string.IsNullOrWhiteSpace(page.Title) ? "当前页面" : page.Title;
				Render(comments);
			}
			catch (Exception exc)
			{
				Logger.Current.WriteLine("cannot refresh embedded comments", exc);
				pageLabel.Text = "读取评论失败（请查看 OneMore 日志）";
			}
			finally
			{
				refreshing.Release();
			}
		}

		private void Render(List<CommentRecord> comments)
		{
			cards.SuspendLayout();
			cards.Controls.Clear();
			if (comments.Count == 0)
			{
				cards.Controls.Add(new Label
				{
					Text = "此页还没有评论。\r\n选中文字后，右键选择“添加评论”。",
					ForeColor = Color.DimGray,
					AutoSize = false,
					Width = 340,
					Height = 70,
					Padding = new Padding(8)
				});
			}
			else
			{
				foreach (var comment in comments.OrderBy(c => c.Resolved).ThenBy(c => c.CreatedUtc))
				{
					cards.Controls.Add(BuildCard(comment));
				}
			}
			cards.ResumeLayout();
		}

		private Control BuildCard(CommentRecord record)
		{
			var card = new Panel
			{
				Width = 345,
				Height = 154,
				Margin = new Padding(0, 0, 0, 10),
				Padding = new Padding(10),
				BackColor = record.Resolved ? Color.FromArgb(239, 246, 239) : Color.White,
				BorderStyle = BorderStyle.FixedSingle
			};
			var status = new Label
			{
				Text = record.Attached ? (record.Resolved ? "✓ 已解决" : "● 已连接") : "⚠ 原文待重连",
				ForeColor = record.Attached ? (record.Resolved ? Color.SeaGreen : Color.DarkGoldenrod) : Color.Firebrick,
				AutoSize = true,
				Location = new Point(9, 8)
			};
			var quote = new Label
			{
				Text = "“" + record.Quote + "”",
				AutoEllipsis = true,
				Location = new Point(9, 34),
				Width = 320,
				Height = 24,
				ForeColor = Color.DimGray
			};
			var body = new Label
			{
				Text = record.Comment,
				AutoEllipsis = true,
				Location = new Point(9, 62),
				Width = 320,
				Height = 42
			};
			var jump = new Button { Text = "跳转", Width = 58, Height = 26, Location = new Point(9, 113) };
			var edit = new Button { Text = "编辑", Width = 58, Height = 26, Location = new Point(73, 113) };
			var resolve = new Button { Text = record.Resolved ? "重开" : "解决", Width = 58, Height = 26, Location = new Point(137, 113) };
			var delete = new Button { Text = "删除", Width = 58, Height = 26, Location = new Point(201, 113) };
			jump.Enabled = record.Attached;
			jump.Click += async (_, _) =>
			{
				await using var one = new OneNote();
				await one.NavigateTo(record.PageId, record.AnchorObjectId ?? string.Empty);
			};
			edit.Click += async (_, _) =>
			{
				using var dialog = new CommentEditDialog(record.Quote, record.Comment);
				if (dialog.ShowDialog(this) == DialogResult.OK)
				{
					await Mutate(record.Id, c => { c.Comment = dialog.CommentText; c.UpdatedUtc = DateTime.UtcNow; });
				}
			};
			resolve.Click += async (_, _) =>
				await Mutate(record.Id, c => { c.Resolved = !c.Resolved; c.UpdatedUtc = DateTime.UtcNow; });
			delete.Click += async (_, _) =>
			{
				if (MessageBox.Show(this, "删除这条评论？", "OneNote 评论",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					await Delete(record.Id);
				}
			};
			card.Controls.Add(status);
			card.Controls.Add(quote);
			card.Controls.Add(body);
			card.Controls.Add(jump);
			card.Controls.Add(edit);
			card.Controls.Add(resolve);
			card.Controls.Add(delete);
			return card;
		}

		private async Task Mutate(string id, Action<CommentRecord> action)
		{
			await using var one = new OneNote();
			var page = await one.GetPage(pageId, OneNote.PageDetail.All);
			var comments = CommentStore.Load(page);
			var record = comments.FirstOrDefault(c => c.Id == id);
			if (record is null) { return; }
			action(record);
			CommentStore.Save(page, comments);
			await one.Update(page);
			await RefreshComments();
		}

		private async Task Delete(string id)
		{
			await using var one = new OneNote();
			var page = await one.GetPage(pageId, OneNote.PageDetail.All);
			var comments = CommentStore.Load(page);
			comments.RemoveAll(c => c.Id == id);
			CommentStore.Save(page, comments);
			await one.Update(page);
			await RefreshComments();
		}
	}


	internal static class PaneNative
	{
		public const int GWL_STYLE = -16;
		public const int WS_CHILD = 0x40000000;
		public const int WS_VISIBLE = 0x10000000;
		public const int WS_POPUP = unchecked((int)0x80000000);
		public const int WS_CAPTION = 0x00C00000;
		public const int WS_THICKFRAME = 0x00040000;
		public const int WS_CLIPSIBLINGS = 0x04000000;
		public const uint SWP_NOZORDER = 0x0004;
		public const uint SWP_NOACTIVATE = 0x0010;
		public const uint SWP_SHOWWINDOW = 0x0040;

		[StructLayout(LayoutKind.Sequential)]
		public struct Rect { public int Left, Top, Right, Bottom; }

		[DllImport("user32.dll", SetLastError = true)]
		public static extern IntPtr SetParent(IntPtr child, IntPtr parent);

		[DllImport("user32.dll")]
		public static extern IntPtr GetParent(IntPtr handle);

		[DllImport("user32.dll", SetLastError = true)]
		public static extern int GetWindowLong(IntPtr handle, int index);

		[DllImport("user32.dll", SetLastError = true)]
		public static extern int SetWindowLong(IntPtr handle, int index, int value);

		[DllImport("user32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool GetClientRect(IntPtr handle, out Rect rect);

		[DllImport("user32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool SetWindowPos(IntPtr handle, IntPtr after,
			int x, int y, int width, int height, uint flags);
	}
}