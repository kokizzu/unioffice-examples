// Copyright 2026 FoxyUtils ehf. All rights reserved.
//
// This example extracts every image from each sheet of an XLSX file via
// the Sheet.Images() helper and writes the raw bytes to disk.

package main

import (
	"fmt"
	"io"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/common/tempstorage"
	"github.com/unidoc/unioffice/v2/spreadsheet"
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
	wb, err := spreadsheet.Open("images.xlsx")
	if err != nil {
		panic(err)
	}
	defer wb.Close()

	for si, sheet := range wb.Sheets() {
		imgs := sheet.Images()
		fmt.Printf("Sheet %d (%q): %d image(s)\n", si+1, sheet.Name(), len(imgs))
		for i, img := range imgs {
			out := fmt.Sprintf("sheet%d_image%d.%s", si+1, i+1, img.Format())
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
