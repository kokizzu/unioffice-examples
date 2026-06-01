// Copyright 2026 FoxyUtils ehf. All rights reserved.
//
// This example extracts every image from each slide of a PPTX file via
// the Slide.Images() helper and writes the raw bytes to disk.

package main

import (
	"fmt"
	"io"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/common/tempstorage"
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
	ppt, err := presentation.Open("image.pptx")
	if err != nil {
		panic(err)
	}
	defer ppt.Close()

	for si, slide := range ppt.Slides() {
		imgs := slide.Images()
		fmt.Printf("Slide %d: %d image(s)\n", si+1, len(imgs))
		for i, img := range imgs {
			out := fmt.Sprintf("slide%d_image%d.%s", si+1, i+1, img.Format())
			if err := saveImage(img.Path(), out); err != nil {
				panic(err)
			}
			sz := img.Size()
			fmt.Printf("  saved %s (%dx%d)\n", out, sz.X, sz.Y)
		}
	}
}

// saveImage copies an image out of unioffice's tempstorage to disk.
func saveImage(src, dst string) error {
	in, err := tempstorage.Open(src)
	if err != nil {
		return err
	}
	defer in.Close()
	out, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer out.Close()
	_, err = io.Copy(out, in)
	return err
}
