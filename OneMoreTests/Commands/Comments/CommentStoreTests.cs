// SPDX-License-Identifier: MPL-2.0
// Copyright 2026 Comments for OneNote contributors

namespace River.OneMoreAddIn.Tests.Commands.Comments
{
	using Microsoft.VisualStudio.TestTools.UnitTesting;
	using River.OneMoreAddIn.Commands;
	using River.OneMoreAddIn.Models;
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Xml.Linq;

	[TestClass]
	public class CommentStoreTests
	{
		private static readonly XNamespace Ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

		[TestMethod]
		public void SaveAndLoadRoundTripsEmbeddedMetadata()
		{
			var page = MakePage(("p1", "before selected words after"));
			var source = new List<CommentRecord>
			{
				new CommentRecord
				{
					Id = "c1", PageId = "page1", AnchorObjectId = "p1",
					Quote = "selected words", Prefix = "before ", Suffix = " after",
					Comment = "review this", CreatedUtc = DateTime.UtcNow,
					UpdatedUtc = DateTime.UtcNow
				}
			};

			CommentStore.Save(page, source);
			var loaded = CommentStore.Load(page);

			Assert.AreEqual(1, loaded.Count);
			Assert.AreEqual("review this", loaded[0].Comment);
			Assert.AreEqual("selected words", loaded[0].Quote);
			StringAssert.Contains(page.GetMetaContent(CommentStore.MetaName), "review this");
		}

		[TestMethod]
		public void RebindFollowsTextWhenParagraphObjectIdChanges()
		{
			var page = MakePage(
				("old", "the original paragraph no longer has it"),
				("new", "before selected words after"));
			var record = new CommentRecord
			{
				AnchorObjectId = "old", Quote = "selected words",
				Prefix = "before ", Suffix = " after"
			};

			var changed = CommentStore.Rebind(page, record);

			Assert.IsTrue(changed);
			Assert.IsTrue(record.Attached);
			Assert.AreEqual("new", record.AnchorObjectId);
		}

		[TestMethod]
		public void RebindDoesNotGuessBetweenAmbiguousQuotesWithoutContext()
		{
			var page = MakePage(("p1", "same quote"), ("p2", "same quote"));
			var record = new CommentRecord { AnchorObjectId = "gone", Quote = "same quote" };

			var changed = CommentStore.Rebind(page, record);

			Assert.IsFalse(changed);
			Assert.IsFalse(record.Attached);
			Assert.AreEqual("gone", record.AnchorObjectId);
		}

		[TestMethod]
		public void SelectionOffsetIdentifiesTheSelectedDuplicate()
		{
			var first = new XElement(Ns + "T", new XCData("same quote then "));
			var selected = new XElement(Ns + "T",
				new XAttribute("selected", "all"), new XCData("same quote"));
			var paragraph = new XElement(Ns + "OE", first, selected);

			var offset = CommentStore.GetSelectionOffset(paragraph, new[] { selected });

			Assert.AreEqual("same quote then ".Length, offset);
		}

		[TestMethod]
		public void RebindChoosesTheExactOccurrenceUsingContext()
		{
			var text = "first: same quote done; second: same quote finished";
			var page = MakePage(("p1", text));
			var expected = text.LastIndexOf("same quote", StringComparison.Ordinal);
			var record = new CommentRecord
			{
				AnchorObjectId = "gone", Quote = "same quote",
				Prefix = "done; second: ", Suffix = " finished"
			};

			var changed = CommentStore.Rebind(page, record);

			Assert.IsTrue(changed);
			Assert.IsTrue(record.Attached);
			Assert.AreEqual("p1", record.AnchorObjectId);
			Assert.AreEqual(expected, record.AnchorOffset);
		}

		[TestMethod]
		public void RebindDoesNotGuessWhenBestContextScoresAreTied()
		{
			var page = MakePage(
				("p1", "before same quote after"),
				("p2", "before same quote after"));
			var record = new CommentRecord
			{
				AnchorObjectId = "gone", Quote = "same quote",
				Prefix = "before ", Suffix = " after"
			};

			var changed = CommentStore.Rebind(page, record);

			Assert.IsFalse(changed);
			Assert.IsFalse(record.Attached);
			Assert.AreEqual("gone", record.AnchorObjectId);
		}

		[TestMethod]
		public void RemoveHighlightClearsOnlyTheDeletedCommentRange()
		{
			const string html = "<span style='background:#FFF2CC'>same quote</span> and " +
				"<span style='font-weight:bold;background:#fff2cc'>same quote</span>";
			var page = MakePage(("p1", html));
			var record = new CommentRecord
			{
				AnchorObjectId = "p1", AnchorOffset = 15,
				HighlightOffset = 15, HighlightLength = 10, Quote = "same quote"
			};

			var changed = CommentStore.RemoveHighlight(page, record);
			var run = page.Root.Descendants(Ns + "T").Single();
			var spans = run.GetCData().GetWrapper().Elements("span").ToList();

			Assert.IsTrue(changed);
			StringAssert.Contains(spans[0].Attribute("style").Value, "background");
			Assert.AreEqual("font-weight:bold", spans[1].Attribute("style").Value);
		}
		[TestMethod]
		public void RemoveHighlightPreservesAUserBackgroundThatThePluginDidNotApply()
		{
			const string html = "<span style='background:#FFF2CC'>same quote</span>";
			var page = MakePage(("p1", html));
			var record = new CommentRecord
			{
				AnchorObjectId = "p1", AnchorOffset = 0, HighlightOffset = 0,
				HighlightLength = 10, HighlightApplied = false, Quote = "same quote"
			};

			var changed = CommentStore.RemoveHighlight(page, record);
			var css = page.Root.Descendants(Ns + "T").Single().GetCData()
				.GetWrapper().Element("span").Attribute("style").Value;

			Assert.IsFalse(changed);
			StringAssert.Contains(css, "background");
		}
		private static Page MakePage(params (string id, string text)[] paragraphs)
		{
			var children = new XElement(Ns + "OEChildren");
			foreach (var paragraph in paragraphs)
			{
				children.Add(new XElement(Ns + "OE",
					new XAttribute("objectID", paragraph.id),
					new XElement(Ns + "T", new XCData(paragraph.text))));
			}

			return new Page(new XElement(Ns + "Page",
				new XAttribute(XNamespace.Xmlns + "one", Ns.NamespaceName),
				new XAttribute("ID", "page1"),
				new XElement(Ns + "Outline", children)));
		}
	}
}