/*
 * Copyright 2026 FoxyUtils ehf. All rights reserved.
 *
 * This example demonstrates structured document tags (Word content controls):
 * template regions with a stable tag name that a program can find and fill in
 * without depending on the surrounding layout.
 */

package main

import (
	"fmt"
	"os"

	"github.com/unidoc/unioffice/v2/color"
	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/document"
	"github.com/unidoc/unioffice/v2/measurement"
	"github.com/unidoc/unioffice/v2/schema/soo/wml"
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

	doc.AddParagraph().AddRun().AddText("Structured document tags (content controls):")

	// Plain text control: a single-line field meant to be filled in. Word
	// shows the placeholder text (grayed out) until the user types over it.
	nameField := doc.AddStructuredDocumentTag()
	nameField.SetTag("customer_name")
	nameField.SetAlias("Customer Name")
	nameField.SetText(false)
	nameField.SetPlaceholder("DefaultPlaceholder")
	nameField.SetShowingPlaceholder(true)
	nameField.SetContentText("Click here to enter the customer's name.")

	// Rich text control: unlike a plain text control, it can hold formatted
	// paragraphs, tables, and images - anything a normal document body can.
	// Locking it keeps the control itself from being deleted from the
	// template while leaving its content editable.
	termsField := doc.AddStructuredDocumentTag()
	termsField.SetTag("terms")
	termsField.SetAlias("Terms And Conditions")
	termsField.SetRichText()
	termsField.SetLock(document.SdtLockSdtLocked)
	run := termsField.AddParagraph().AddRun()
	run.AddText("Payment is due within 30 days of the invoice date.")
	run.Properties().SetItalic(true)

	// Date picker control, with a display format Word uses to render the
	// stored date.
	dateField := doc.AddStructuredDocumentTag()
	dateField.SetTag("invoice_date")
	dateField.SetAlias("Invoice Date")
	dateField.SetDate("M/d/yyyy")
	dateField.SetContentText("1/1/2026")

	// Drop-down list control: the user can only pick one of the listed
	// display values; the document stores the matching Value, not the text
	// shown on screen.
	statusField := doc.AddStructuredDocumentTag()
	statusField.SetTag("status")
	statusField.SetAlias("Status")
	statusField.SetDropDownList(
		document.SdtListItem{DisplayText: "Draft", Value: "draft"},
		document.SdtListItem{DisplayText: "Sent", Value: "sent"},
		document.SdtListItem{DisplayText: "Paid", Value: "paid"},
	)
	statusField.SetContentText("Draft")

	// Combo box control: like a drop-down, but also accepts free-form typed
	// text instead of only the listed items.
	regionField := doc.AddStructuredDocumentTag()
	regionField.SetTag("region")
	regionField.SetAlias("Sales Region")
	regionField.SetComboBox(
		document.SdtListItem{DisplayText: "North America", Value: "na"},
		document.SdtListItem{DisplayText: "Europe", Value: "eu"},
		document.SdtListItem{DisplayText: "Asia Pacific", Value: "apac"},
	)
	regionField.SetContentText("North America")

	// Rich text control containing a table. Word only allows tables (and
	// other block content) inside rich-text or group controls, not inside
	// plain text, date, combo box, or drop-down controls.
	itemsField := doc.AddStructuredDocumentTag()
	itemsField.SetTag("order_items")
	itemsField.SetAlias("Order Items")
	itemsField.SetRichText()

	itemsTable := itemsField.AddTable()
	itemsTable.Properties().SetWidthPercent(100)
	itemsTable.Properties().Borders().SetAll(wml.ST_BorderSingle, color.Auto, measurement.Point)

	header := itemsTable.AddRow()
	header.AddCell().AddParagraph().AddRun().AddText("Item")
	header.AddCell().AddParagraph().AddRun().AddText("Quantity")

	row := itemsTable.AddRow()
	row.AddCell().AddParagraph().AddRun().AddText("Widget")
	row.AddCell().AddParagraph().AddRun().AddText("4")

	row = itemsTable.AddRow()
	row.AddCell().AddParagraph().AddRun().AddText("Gadget")
	row.AddCell().AddParagraph().AddRun().AddText("2")

	// Rich text control containing a bulleted list, built the same way as a
	// bulleted list in the document body: a numbering definition with a
	// bullet-format level, applied to each list paragraph.
	optionsField := doc.AddStructuredDocumentTag()
	optionsField.SetTag("delivery_options")
	optionsField.SetAlias("Delivery Options")
	optionsField.SetRichText()

	bulletDef := doc.Numbering.AddDefinition()
	bulletLvl := bulletDef.AddLevel()
	bulletLvl.SetFormat(wml.ST_NumberFormatBullet)
	bulletLvl.SetAlignment(wml.ST_JcLeft)
	bulletLvl.RunProperties().SetFontFamily("Symbol")
	bulletLvl.SetText("")
	bulletLvl.Properties().SetLeftIndent(0.5 * measurement.Inch)

	for _, option := range []string{"Standard shipping", "Express shipping", "Local pickup"} {
		p := optionsField.AddParagraph()
		p.SetNumberingDefinition(bulletDef)
		p.SetNumberingLevel(0)
		p.AddRun().AddText(option)
	}

	// Inline (run-level) controls live inside a paragraph next to regular
	// runs, rather than as their own block - useful for a short field
	// embedded mid-sentence.
	summary := doc.AddParagraph()
	summary.AddRun().AddText("Invoice status: ")
	inlineStatus := summary.AddStructuredDocumentTag()
	inlineStatus.SetTag("status_inline")
	inlineStatus.SetText(false)
	inlineStatus.SetContentText("Draft")
	summary.AddRun().AddText(".")

	if err := doc.SaveToFile("structured-document-tags.docx"); err != nil {
		panic(err)
	}

	// StructuredDocumentTags finds every block-level control in the document,
	// descending into tables, headers, and footers - letting a program
	// discover and fill in a template's fields without knowing its layout.
	fmt.Println("Block-level structured document tags:")
	for _, sdt := range doc.StructuredDocumentTags() {
		fmt.Printf("- tag=%q alias=%q type=%v text=%q\n", sdt.Tag(), sdt.Alias(), sdt.Type(), sdt.Text())
	}

	// Inline controls are scoped to the paragraph they live in.
	fmt.Println("\nInline structured document tags:")
	for _, sdt := range summary.StructuredDocumentTags() {
		fmt.Printf("- tag=%q type=%v text=%q\n", sdt.Tag(), sdt.Type(), sdt.Text())
	}
}
