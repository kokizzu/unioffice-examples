/*
 * Copyright 2026 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates the bookmark API: adding named anchors,
 * targeting them with hyperlinks, enumerating and looking them up, and
 * editing the paragraph that holds an anchor.
 */

package main

import (
	"fmt"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/document"
)

func init() {
	// Make sure to load your metered License API key prior to using the library.
	// If you need a key, you can sign up and create a free one at https://cloud.unidoc.io
	err := license.SetMeteredKey(os.Getenv(`UNIDOC_LICENSE_API_KEY`))
	if err != nil {
		panic(err)
	}
}

func main() {
	doc := document.New()
	defer doc.Close()

	title := doc.AddParagraph()
	title.SetStyle("Title")
	title.AddRun().AddText("Bookmarks API")

	// Anchor a paragraph by name.
	intro := doc.AddParagraph()
	intro.AddBookmark("intro")
	intro.AddRun().AddText("Introductory paragraph, anchored by \"intro\".")

	// Anchor a heading. AddBookmark allocates a fresh ID for each call.
	h1 := doc.AddParagraph()
	h1.SetStyle("Heading1")
	h1.AddBookmark("section1")
	h1.AddRun().AddText("Section One")

	body := doc.AddParagraph()
	body.AddRun().AddText("Body text follows the heading anchor.")

	summary := doc.AddParagraph()
	summary.AddBookmark("summary")
	summary.AddRun().AddText("Summary: placeholder value")

	// Hyperlink targeting a bookmark by reference.
	link := doc.AddParagraph()
	hl := link.AddHyperLink()
	section1, _ := doc.BookmarkByName("section1")
	hl.SetTargetBookmark(section1)
	hl.AddRun().AddText("Jump to Section One")

	// Enumerate every bookmark in the document.
	fmt.Println("Bookmarks:")
	for _, b := range doc.Bookmarks() {
		shape := "range"
		if b.IsEmpty() {
			shape = "anchor"
		}
		fmt.Printf("  - %-10s id=%d %s text=%q\n", b.Name(), b.ID(), shape, b.Text())
	}

	// AddBookmark creates a zero-width anchor, so use Paragraphs() to reach
	// the paragraph it sits in and append text to it.
	bm, _ := doc.BookmarkByName("summary")
	p := bm.Paragraphs()[0]
	p.AddRun().AddText(" -- appended via bookmark lookup.")

	if err := doc.SaveToFile("bookmarks.docx"); err != nil {
		panic(err)
	}
}
