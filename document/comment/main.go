/*
 * Copyright 2025 FoxyUtils ehf. All rights reserved.
 *
 * This example showcases how to add comment to paragraph in document file.
 */

package main

import (
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/document"
	"github.com/unidoc/unioffice/v2/measurement"
)

func init() {
	// This example requires both for unioffice and unipdf.
	err := license.SetMeteredKey(os.Getenv(`UNIDOC_LICENSE_API_KEY`))
	if err != nil {
		panic(err)
	}
}

func main() {
	doc := document.New()
	defer doc.Close()

	para := doc.AddParagraph()
	para.SetAfterSpacing(measurement.Distance(20))

	// Initiate a comment block with the author name and comment text that will enclosed the defined paragraph runs.
	cmId := para.AddComment("UniOffice User", "This is comment")

	run := para.AddRun()
	run.AddText("Lorem")

	// Close the comment block.
	para.CloseComment(cmId)

	// Reply to the comment, building a thread. Replies must be added after the
	// comment they reply to has been closed (CloseComment), since a reply's
	// markers are anchored right next to the thread root's.
	firstComment := doc.Comment(cmId)
	reply, err := firstComment.AddReply("Reviewer One", "I agree, this should be simplified.")
	if err != nil {
		panic(err)
	}

	// Replying to a reply still attaches to the thread root - Word threads are
	// flat, so this appears as a second reply alongside the first.
	if _, err := reply.AddReply("UniOffice User", "Good point, updating it now."); err != nil {
		panic(err)
	}

	// Mark the thread as resolved once the discussion is settled.
	firstComment.SetDone(true)

	// This following paragraph will not have any comment since the comment block is closed.
	run = para.AddRun()
	run.AddText(" ipsum dolor sit amet consectetur adipiscing elit, urna consequat felis vehicula class ultricies mollis dictumst, aenean non a in donec nulla.")

	para = doc.AddParagraph()
	para.SetAfterSpacing(measurement.Distance(20))
	run = para.AddRun()
	run.AddText("Phasellus ante pellentesque erat cum risus consequat imperdiet aliquam, ")

	// Initiaate a new comment block that will enclose the following paragraph runs and the next paragraph runs.
	cmId = para.AddComment("Other User", "Second comment")

	run = para.AddRun()
	run.AddText("integer placerat et turpis mi eros nec lobortis taciti, vehicula nisl litora tellus ligula porttitor metus.")

	para = doc.AddParagraph()
	para.SetAfterSpacing(measurement.Distance(20))
	run = para.AddRun()
	run.AddText("Dapibus imperdiet praesent magnis ridiculus congue gravida curabitur dictum sagittis, enim et magna sit inceptos sodales parturient pharetra mollis, aenean vel nostra tellus commodo pretium sapien sociosqu.")

	// close the comment block.
	para.CloseComment(cmId)

	doc.SaveToFile("simple_doc_with_comment.docx")
}
