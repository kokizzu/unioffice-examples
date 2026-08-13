/*
 * Copyright 2025 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates how to retrieve and display comments stored within a DOCX file,
 * including reply threads.
 */

package main

import (
	"fmt"
	"log"
	"os"
	"strings"

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

	// Check if the document has comments.
	if !doc.HasComments() {
		fmt.Println("Document has no comment")
		return
	}

	comments := doc.Comments()

	fmt.Printf("Document has %d comments:\n\n", len(comments))

	// Print each top level comment followed by its replies, indented - the
	// same threading Word shows in its Comments pane. Comments.IsReply()
	// tells apart thread roots from replies, and Comment.Replies() walks a
	// thread in display order.
	for _, c := range comments {
		if c.IsReply() {
			continue
		}
		printThread(c, 0)
	}
}

func printThread(c document.Comment, depth int) {
	indent := strings.Repeat("    ", depth)
	resolved := ""
	if c.Done() {
		resolved = " [resolved]"
	}
	fmt.Printf("%s%d. Comment by %s: %s%s\n", indent, c.ID(), c.Author(), c.Text(), resolved)

	for _, reply := range c.Replies() {
		printThread(reply, depth+1)
	}
}
