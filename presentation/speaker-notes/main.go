/*
 * Copyright 2026 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates speaker notes: per-slide presenter-only text that
 * appears in PowerPoint's Notes pane and presenter view, not on the slide
 * itself.
 */

package main

import (
	"fmt"
	"log"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/presentation"
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
	ppt := presentation.New()
	defer ppt.Close()

	intro := ppt.AddSlide()
	intro.AddTextBox().AddParagraph().AddRun().SetText("Welcome")
	// EnsureNotes creates the notes part (and a notes master, if the
	// presentation has none yet) on first use, and is safe to call again -
	// it returns the same notes every time.
	intro.EnsureNotes().SetText("Greet the audience and introduce the topic before advancing.")

	agenda := ppt.AddSlide()
	agenda.AddTextBox().AddParagraph().AddRun().SetText("Agenda")
	// A multi-paragraph note: SetText replaces the notes with a single
	// paragraph, and AddParagraph appends one more line after it.
	agendaNotes := agenda.EnsureNotes()
	agendaNotes.SetText("Cover three points today:")
	agendaNotes.AddParagraph().AddRun().SetText("1. What changed")
	agendaNotes.AddParagraph().AddRun().SetText("2. Why it matters")
	agendaNotes.AddParagraph().AddRun().SetText("3. What's next")

	// This slide is left without notes, to show GetNotes reporting none below.
	closing := ppt.AddSlide()
	closing.AddTextBox().AddParagraph().AddRun().SetText("Questions?")

	if err := ppt.Validate(); err != nil {
		log.Fatal(err)
	}
	ppt.SaveToFile("speaker-notes.pptx")

	// GetNotes reports whether a slide has notes at all, without creating
	// them - unlike EnsureNotes, it never modifies the slide.
	for i, slide := range ppt.Slides() {
		notes, ok := slide.GetNotes()
		if !ok {
			fmt.Printf("Slide %d: no notes\n", i+1)
			continue
		}
		fmt.Printf("Slide %d notes:\n%s\n\n", i+1, notes.Text())
	}
}
