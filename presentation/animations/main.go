/*
 * Copyright 2026 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates click-triggered entrance animations on slide
 * shapes: each click during the slide show reveals or animates in one more
 * shape, instead of showing everything on the slide at once.
 */

package main

import (
	"log"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/measurement"
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

	slide := ppt.AddSlide()

	title := slide.AddTextBox()
	title.AddParagraph().AddRun().SetText("Quarterly Highlights")
	title.Properties().SetPosition(0.5*measurement.Inch, 0.5*measurement.Inch)
	title.Properties().SetWidth(6 * measurement.Inch)

	// Animations returns the slide's main click-driven animation sequence,
	// creating the timing scaffold on first use. Every AddOnClickEntrance
	// call below appends one more click step to it.
	seq := slide.Animations()

	// Reveal the bullet points one at a time: each gets its own entrance, so
	// the next one only appears on the next click instead of all at once.
	bullets := []string{
		"Revenue is up 12% year over year",
		"Two new markets launched",
		"Support tickets down 30%",
	}
	for i, text := range bullets {
		tb := slide.AddTextBox()
		tb.AddParagraph().AddRun().SetText(text)
		tb.Properties().SetPosition(0.5*measurement.Inch, measurement.Distance(2+i)*measurement.Inch)
		tb.Properties().SetWidth(6 * measurement.Inch)

		seq.AddOnClickEntrance(tb, presentation.EntranceAppear)
	}

	// The other entrance effects animate the shape in rather than just
	// revealing it: EntranceFade brings it in via opacity, EntranceFlyIn
	// slides it in from the bottom of the slide.
	callout := slide.AddTextBox()
	callout.AddParagraph().AddRun().SetText("Ask us about the new pricing tiers")
	callout.Properties().SetPosition(0.5*measurement.Inch, 5.5*measurement.Inch)
	callout.Properties().SetWidth(6 * measurement.Inch)
	seq.AddOnClickEntrance(callout, presentation.EntranceFlyIn)

	closing := slide.AddTextBox()
	closing.AddParagraph().AddRun().SetText("Thank you")
	closing.Properties().SetPosition(0.5*measurement.Inch, 6.5*measurement.Inch)
	closing.Properties().SetWidth(6 * measurement.Inch)
	seq.AddOnClickEntrance(closing, presentation.EntranceFade)

	if err := ppt.Validate(); err != nil {
		log.Fatal(err)
	}
	ppt.SaveToFile("animations.pptx")
}
