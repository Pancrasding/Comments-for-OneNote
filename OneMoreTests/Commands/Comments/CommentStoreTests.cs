// SPDX-License-Identifier: MPL-2.0
// Copyright 2026 OneMore Comments contributors

namespace River.OneMoreAddIn.Tests.Commands.Comments
{
	using Microsoft.VisualStudio.TestTools.UnitTesting;
	using River.OneMoreAddIn.Commands;
	using River.OneMoreAddIn.Models;
	using System;
	using System.Collections.Generic;
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