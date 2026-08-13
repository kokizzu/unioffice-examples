/*
 * Copyright 2025 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates how to remove a comment from a DOCX file. Removing
 * a comment that is the root of a reply thread removes the whole thread -
 * Word never leaves a dangling reply behind.
 */

package main

import (
	"fmt"
	"log"
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
	doc, err := document.Open("sample.docx")
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}
	defer doc.Close()

	listComments(doc)

	// Find the first top-level comment (a thread root, not a reply) rather
	// than hard-coding an ID, so the example works against any document, not
	// just the bundled sample.
	root, ok := firstThreadRoot(doc)
	if !ok {
		fmt.Println("No comments to remove")
		return
	}

	// Removing a thread root takes any replies in its thread with it.
	if ok := doc.RemoveComment(root.ID()); !ok {
		fmt.Println("Failed removing comment")
		return
	}

	fmt.Println("\nComment removed successfully, along with any replies in its thread.")
	fmt.Println("")

	listComments(doc)
}

// firstThreadRoot returns the first top-level comment in document order (one
// that is not itself a reply), or ok=false if the document has no comments.
func firstThreadRoot(doc *document.Document) (comment document.Comment, ok bool) {
	for _, c := range doc.Comments() {
		if !c.IsReply() {
			return c, true
		}
	}
	return document.Comment{}, false
}

func listComments(doc *document.Document) {
	comments := doc.Comments()
	fmt.Printf("Document has %d comments.\n", len(comments))

	for _, c := range comments {
		marker := ""
		if c.IsReply() {
			marker = " (reply)"
		}
		fmt.Printf("%d. Comment by %s%s: %s\n", c.ID(), c.Author(), marker, c.Text())
	}
}
