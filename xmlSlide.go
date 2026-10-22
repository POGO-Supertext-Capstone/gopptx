// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"encoding/xml"
	"sync"
)

type Slide struct {
	mu                     sync.Mutex
	XMLName                xml.Name          `xml:"p:sld"`
	XMLNSA                 string            `xml:"xmlns:a,attr"`
	XMLNSP                 string            `xml:"xmlns:p,attr"`
	XMLNSR                 string            `xml:"xmlns:r,attr"`
	XMLNSP14               string            `xml:"xmlns:p14,attr"`
	XMLNSP15               string            `xml:"xmlns:p15,attr"`
	XMLNSMC                string            `xml:"xmlns:mc,attr"`
	CommonSlideData        SlideData         `xml:"p:cSld"`
	ColorMapOverride       *ClrMapOverride   `xml:"p:clrMapOvr,omitempty"`
	AlternateContent       *alternateContent `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML         `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

type SlideData struct {
	Background *bg       `xml:"p:bg,omitempty"`
	ShapeTree  ShapeTree `xml:"p:spTree"`
}

// bg defines the slide background (p:bg) element of the common slide data.
type bg struct {
	BGProperties *bgProperties `xml:"p:bgPr,omitempty"`
}

// bgProperties defines the background properties (p:bgPr) element.
type bgProperties struct {
	NoFill     *noFill    `xml:"a:noFill,omitempty"`
	SolidFill  *SolidFill `xml:"a:solidFill,omitempty"`
	EffectList *struct{}  `xml:"a:effectLst,omitempty"`
}

type ShapeTree struct {
	NonVisualGroupShapeProperties *NonVisualGroupShapeProperties `xml:"p:nvGrpSpPr,omitempty"`
	GroupShapeProperties          *GroupShapeProperties          `xml:"p:grpSpPr,omitempty"`
	Shape                         []Shape                        `xml:"p:sp"`
	Picture                       []Picture                      `xml:"p:pic"`
}

type NonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *CommonNonVisualProperties           `xml:"p:cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"p:cNvGrpSpPr"`
	NonVisualProperties                 *NonVisualProperties                 `xml:"p:nvPr"`
}

type GroupShapeProperties struct {
	Xfrm *Xfrm `xml:"a:xfrm"`
}

type Xfrm struct {
	Offset       *Offset  `xml:"a:off"`
	Extents      *Extents `xml:"a:ext"`
	ChildOffset  *Offset  `xml:"a:chOff"`
	ChildExtents *Extents `xml:"a:chExt"`
}

type Shape struct {
	NonVisualShapeProperties *NonVisualShapeProperties `xml:"p:nvSpPr"`
	ShapeProperties          *ShapeProperties          `xml:"p:spPr"`
	TextBody                 *TextBody                 `xml:"p:txBody,omitempty"`
}

type NonVisualShapeProperties struct {
	CommonNonVisualProperties      *CommonNonVisualProperties      `xml:"p:cNvPr"`
	CommonNonVisualShapeProperties *CommonNonVisualShapeProperties `xml:"p:cNvSpPr"`
	NonVisualProperties            *NonVisualProperties            `xml:"p:nvPr"`
}

type NonVisualProperties struct {
	Ph *Ph `xml:"p:ph,omitempty"`
}

// Picture defines a picture shape (p:pic) in the slide shape tree.
type Picture struct {
	NonVisualPictureProperties *NonVisualPictureProperties `xml:"p:nvPicPr"`
	BlipFill                   *BlipFill                   `xml:"p:blipFill"`
	ShapeProperties            *ShapeProperties            `xml:"p:spPr"`
}

// NonVisualPictureProperties defines the non-visual properties of a picture
// (p:nvPicPr).
type NonVisualPictureProperties struct {
	CommonNonVisualProperties        *CommonNonVisualProperties        `xml:"p:cNvPr"`
	CommonNonVisualPictureProperties *CommonNonVisualPictureProperties `xml:"p:cNvPicPr"`
	NonVisualProperties              *NonVisualProperties              `xml:"p:nvPr"`
}

// CommonNonVisualPictureProperties defines the common non-visual properties
// of a picture (p:cNvPicPr).
type CommonNonVisualPictureProperties struct {
	PreferRelativeResize *int `xml:"preferRelativeResize,attr,omitempty"`
}

// BlipFill defines the blip fill (p:blipFill) element of a picture.
type BlipFill struct {
	Blip    *Blip    `xml:"a:blip"`
	Stretch *Stretch `xml:"a:stretch"`
}

// Blip defines the blip (a:blip) element of the blip fill. The image part is
// referenced by the relationship id in the embed attribute.
type Blip struct {
	Embed       string       `xml:"r:embed,attr"`
	AlphaModFix *AlphaModFix `xml:"a:alphaModFix,omitempty"`
}

// AlphaModFix defines the alpha modulation fix (a:alphaModFix) element of
// the blip.
type AlphaModFix struct {
	Amount *int `xml:"amt,attr,omitempty"`
}

// Stretch defines the stretch (a:stretch) element of the blip fill.
type Stretch struct {
	FillRect *FillRect `xml:"a:fillRect"`
}

// FillRect defines the fill rectangle (a:fillRect) element of the stretch.
type FillRect struct {
	Left  *int `xml:"l,attr,omitempty"`
	Top   *int `xml:"t,attr,omitempty"`
	Right *int `xml:"r,attr,omitempty"`
	Bot   *int `xml:"b,attr,omitempty"`
}

// ClrMapOverride defines the color map override (p:clrMapOvr) element of
// the slide.
type ClrMapOverride struct {
	MasterClrMapping *MasterClrMapping `xml:"a:masterClrMapping"`
}

// MasterClrMapping defines the master color mapping (a:masterClrMapping)
// element.
type MasterClrMapping struct{}

type CommonNonVisualShapeProperties struct {
	TxBox      *XMLBool    `xml:"txBox,attr,omitempty"`
	ShapeLocks *ShapeLocks `xml:"a:spLocks,omitempty"`
}

type ShapeProperties struct {
	Xfrm           *Xfrm           `xml:"a:xfrm"`
	PresetGeometry *PresetGeometry `xml:"a:prstGeom,omitempty"`
	NoFill         *noFill         `xml:"a:noFill,omitempty"`
	SolidFill      *SolidFill      `xml:"a:solidFill,omitempty"`
	Ln             *Line           `xml:"a:ln,omitempty"`
}

type PresetGeometry struct {
	Preset          string           `xml:"prst,attr"`
	AdjustValueList *AdjustValueList `xml:"a:avLst"`
}

type Line struct {
	Width     *int       `xml:"w,attr,omitempty"`
	NoFill    *noFill    `xml:"a:noFill,omitempty"`
	SolidFill *SolidFill `xml:"a:solidFill,omitempty"`
}

type TextBody struct {
	BodyProperties *BodyProperties `xml:"a:bodyPr"`
	ListStyle      *struct{}       `xml:"a:lstStyle,omitempty"`
	Paragraph      []Paragraph     `xml:"a:p"`
}

type BodyProperties struct {
	AnchorCtr        *int         `xml:"anchorCtr,attr,omitempty"`
	Anchor           *string      `xml:"anchor,attr,omitempty"`
	BIns             *int         `xml:"bIns,attr,omitempty"`
	LIns             *int         `xml:"lIns,attr,omitempty"`
	SpcFirstLastPara *int         `xml:"spcFirstLastPara,attr,omitempty"`
	RIns             *int         `xml:"rIns,attr,omitempty"`
	Wrap             *string      `xml:"wrap,attr,omitempty"`
	TIns             *int         `xml:"tIns,attr,omitempty"`
	NormAutofit      *NormAutofit `xml:"a:normAutofit,omitempty"`
	NoAutofit        *NoAutofit   `xml:"a:noAutofit,omitempty"`
}

// NormAutofit defines the normal autofit (a:normAutofit) element of the body
// properties.
type NormAutofit struct {
	FontScale            *int `xml:"fontScale,attr,omitempty"`
	LineSpacingReduction *int `xml:"lnSpcReduction,attr,omitempty"`
}

type Paragraph struct {
	ParagraphProperties       *ParagraphProperties `xml:"a:pPr,omitempty"`
	Runs                      []Runs               `xml:"a:r"`
	EndParagraphRunProperties *RunProperties       `xml:"a:endParaRPr,omitempty"`
}

// ParagraphProperties defines the paragraph properties (a:pPr) element of
// the paragraph.
type ParagraphProperties struct {
	Indent           *int           `xml:"indent,attr,omitempty"`
	Lvl              *int           `xml:"lvl,attr,omitempty"`
	MarL             *int           `xml:"marL,attr,omitempty"`
	RTL              *int           `xml:"rtl,attr,omitempty"`
	Align            *string        `xml:"algn,attr,omitempty"`
	LineSpacing      *LineSpacing   `xml:"a:lnSpc,omitempty"`
	SpacingBefore    *SpacingChange `xml:"a:spcBef,omitempty"`
	SpacingAfter     *SpacingChange `xml:"a:spcAft,omitempty"`
	BuNone           *struct{}      `xml:"a:buNone,omitempty"`
	DefRunProperties *RunProperties `xml:"a:defRPr,omitempty"`
}

// SpacingChange defines the spacing before (a:spcBef) or spacing after
// (a:spcAft) element of the paragraph properties.
type SpacingChange struct {
	SpacingPercent *SpacingPercent `xml:"a:spcPct,omitempty"`
	SpacingPts     *SpacingPts     `xml:"a:spcPts,omitempty"`
}

type Runs struct {
	RunProperties *RunProperties `xml:"a:rPr,omitempty"`
	Text          string         `xml:"a:t"`
}

type RunProperties struct {
	Bold          *int       `xml:"b,attr,omitempty"`
	Lang          string     `xml:"lang,attr,omitempty"`
	Size          *int       `xml:"sz,attr,omitempty"`
	Space         *int       `xml:"spc,attr,omitempty"`
	Strike        string     `xml:"strike,attr,omitempty"`
	SolidFill     *SolidFill `xml:"a:solidFill,omitempty"`
	Latin         *Font      `xml:"a:latin,omitempty"`
	EastAsian     *Font      `xml:"a:ea,omitempty"`
	ComplexScript *Font      `xml:"a:cs,omitempty"`
}
type SolidFill struct {
	SolidRGBColor *SolidRGBColor `xml:"a:srgbClr"`
}

type decodeSlide struct {
	mu                     sync.Mutex
	XMLName                xml.Name              `xml:"sld"`
	CommonSlideData        decodeSlideData       `xml:"cSld"`
	ColorMapOverride       *decodeClrMapOverride `xml:"clrMapOvr,omitempty"`
	AlternateContent       *alternateContent     `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML             `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

// decodeClrMapOverride directly maps the clrMapOvr element of the slide.
type decodeClrMapOverride struct {
	MasterClrMapping *MasterClrMapping `xml:"masterClrMapping,omitempty"`
}

type decodeSlideData struct {
	Background *DecodeBackground `xml:"bg,omitempty"`
	ShapeTree  decodeShapeTree   `xml:"spTree"`
}

// DecodeBackground directly maps the bg element of the slide.
type DecodeBackground struct {
	BGProperties *DecodeBGProperties `xml:"bgPr,omitempty"`
}

// DecodeBGProperties directly maps the bgPr element of the slide background.
type DecodeBGProperties struct {
	NoFill     *noFill          `xml:"noFill,omitempty"`
	SolidFill  *DecodeSolidFill `xml:"solidFill,omitempty"`
	EffectList *struct{}        `xml:"effectLst,omitempty"`
}

type decodeShapeTree struct {
	NonVisualGroupShapeProperties *decodeNonVisualGroupShapeProperties `xml:"nvGrpSpPr,omitempty"`
	GroupShapeProperties          *decodeGroupShapeProperties          `xml:"grpSpPr,omitempty"`
	Shape                         []decodeShape                        `xml:"sp"`
	Picture                       []decodePicture                      `xml:"pic"`
}

type decodeNonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *CommonNonVisualProperties           `xml:"cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"cNvGrpSpPr"`
	NonVisualProperties                 *decodeNonVisualProperties           `xml:"nvPr"`
}

type CommonNonVisualProperties struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Title string `xml:"title,attr,omitempty"`
}

type CommonNonVisualGroupShapeProperties struct{}

type decodeNonVisualProperties struct {
	Ph *Ph `xml:"ph,omitempty"`
}

type Ph struct {
	Idx  *int    `xml:"idx,attr,omitempty"`
	Type *string `xml:"type,attr,omitempty"`
}

type decodeGroupShapeProperties struct {
	Xfrm *DecodeXfrm `xml:"xfrm"`
}

type DecodeXfrm struct {
	Offset       *Offset  `xml:"off"`
	Extents      *Extents `xml:"ext"`
	ChildOffset  *Offset  `xml:"chOff"`
	ChildExtents *Extents `xml:"chExt"`
}

type Offset struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

type Extents struct {
	CX int `xml:"cx,attr"`
	CY int `xml:"cy,attr"`
}

type decodeShape struct {
	NonVisualShapeProperties *decodeNonVisualShapeProperties `xml:"nvSpPr"`
	ShapeProperties          *DecodeShapeProperties          `xml:"spPr"`
	TextBody                 *DecodeTextBody                 `xml:"txBody,omitempty"`
}

// decodePicture directly maps the pic element in the shape tree.
type decodePicture struct {
	NonVisualPictureProperties *decodeNonVisualPictureProperties `xml:"nvPicPr"`
	BlipFill                   *decodeBlipFill                   `xml:"blipFill"`
	ShapeProperties            *DecodeShapeProperties            `xml:"spPr"`
}

// decodeNonVisualPictureProperties directly maps the nvPicPr element of the
// picture.
type decodeNonVisualPictureProperties struct {
	CommonNonVisualProperties        *CommonNonVisualProperties        `xml:"cNvPr"`
	CommonNonVisualPictureProperties *CommonNonVisualPictureProperties `xml:"cNvPicPr"`
	NonVisualProperties              *decodeNonVisualProperties        `xml:"nvPr"`
}

// decodeBlipFill directly maps the blipFill element of the picture.
type decodeBlipFill struct {
	Blip    *decodeBlip    `xml:"blip"`
	Stretch *decodeStretch `xml:"stretch"`
}

// decodeBlip directly maps the blip element of the blip fill.
type decodeBlip struct {
	Embed       string       `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
	AlphaModFix *AlphaModFix `xml:"alphaModFix,omitempty"`
}

// decodeStretch directly maps the stretch element of the blip fill.
type decodeStretch struct {
	FillRect *FillRect `xml:"fillRect"`
}

type decodeNonVisualShapeProperties struct {
	CommonNonVisualProperties      *CommonNonVisualProperties            `xml:"cNvPr"`
	CommonNonVisualShapeProperties *decodeCommonNonVisualShapeProperties `xml:"cNvSpPr"`
	NonVisualProperties            *decodeNonVisualProperties            `xml:"nvPr"`
}

type decodeCommonNonVisualShapeProperties struct {
	TxBox      *bool       `xml:"txBox,attr,omitempty"`
	ShapeLocks *ShapeLocks `xml:"spLocks,omitempty"`
}

type ShapeLocks struct {
	NoGroup *int `xml:"noGrp,attr,omitempty"`
}

type DecodeShapeProperties struct {
	Xfrm           *DecodeXfrm           `xml:"xfrm"`
	PresetGeometry *DecodePresetGeometry `xml:"prstGeom,omitempty"`
	NoFill         *noFill               `xml:"noFill,omitempty"`
	SolidFill      *DecodeSolidFill      `xml:"solidFill,omitempty"`
	Ln             *decodeLine           `xml:"ln,omitempty"`
}

type noFill struct{}

type DecodePresetGeometry struct {
	Preset          string           `xml:"prst,attr"`
	AdjustValueList *AdjustValueList `xml:"avLst"`
}

type AdjustValueList struct {
}

type AdjustValue struct {
	Name    string `xml:"name,attr"`
	Formula string `xml:"fmla,attr"`
}

type decodeLine struct {
	Width     *int             `xml:"w,attr,omitempty"`
	NoFill    *noFill          `xml:"noFill,omitempty"`
	SolidFill *DecodeSolidFill `xml:"solidFill,omitempty"`
}

type DecodeTextBody struct {
	BodyProperties *DecodeBodyProperties `xml:"bodyPr"`
	ListStyle      *struct{}             `xml:"lstStyle,omitempty"`
	Paragraph      []DecodeParagraph     `xml:"p"`
}

type DecodeBodyProperties struct {
	AnchorCtr        *int         `xml:"anchorCtr,attr,omitempty"`
	Anchor           *string      `xml:"anchor,attr,omitempty"`
	BIns             *int         `xml:"bIns,attr,omitempty"`
	LIns             *int         `xml:"lIns,attr,omitempty"`
	SpcFirstLastPara *int         `xml:"spcFirstLastPara,attr,omitempty"`
	RIns             *int         `xml:"rIns,attr,omitempty"`
	Wrap             *string      `xml:"wrap,attr,omitempty"`
	TIns             *int         `xml:"tIns,attr,omitempty"`
	NormAutofit      *NormAutofit `xml:"normAutofit,omitempty"`
	NoAutofit        *NoAutofit   `xml:"noAutofit,omitempty"`
}

type NoAutofit struct{}

type DecodeParagraph struct {
	ParagraphProperties       *DecodeParagraphProperties `xml:"pPr,omitempty"`
	Runs                      []DecodeRuns               `xml:"r"`
	EndParagraphRunProperties *DecodeRunProperties       `xml:"endParaRPr,omitempty"`
}

// DecodeParagraphProperties directly maps the pPr element of the paragraph.
type DecodeParagraphProperties struct {
	Indent           *int                 `xml:"indent,attr,omitempty"`
	Lvl              *int                 `xml:"lvl,attr,omitempty"`
	MarL             *int                 `xml:"marL,attr,omitempty"`
	RTL              *int                 `xml:"rtl,attr,omitempty"`
	Align            *string              `xml:"algn,attr,omitempty"`
	LineSpacing      *DecodeLineSpacing   `xml:"lnSpc,omitempty"`
	SpacingBefore    *DecodeSpacingChange `xml:"spcBef,omitempty"`
	SpacingAfter     *DecodeSpacingChange `xml:"spcAft,omitempty"`
	BuNone           *struct{}            `xml:"buNone,omitempty"`
	DefRunProperties *DecodeRunProperties `xml:"defRPr,omitempty"`
}

// DecodeLineSpacing directly maps the lnSpc element of the paragraph
// properties.
type DecodeLineSpacing struct {
	SpacingPercent *SpacingPercent `xml:"spcPct,omitempty"`
}

// DecodeSpacingChange directly maps the spcBef or spcAft element of the
// paragraph properties.
type DecodeSpacingChange struct {
	SpacingPercent *SpacingPercent `xml:"spcPct,omitempty"`
	SpacingPts     *SpacingPts     `xml:"spcPts,omitempty"`
}

// SpacingPts defines the spacing in points (a:spcPts) element.
type SpacingPts struct {
	Val int `xml:"val,attr"`
}

type LineSpacing struct {
	SpacingPercent *SpacingPercent `xml:"a:spcPct"`
}

type SpacingPercent struct {
	Val int `xml:"val,attr"`
}

type DecodeRuns struct {
	RunProperties *DecodeRunProperties `xml:"rPr,omitempty"`
	Text          string               `xml:"t"`
}

type DecodeRunProperties struct {
	Bold          *int             `xml:"b,attr,omitempty"`
	Lang          string           `xml:"lang,attr,omitempty"`
	Size          *int             `xml:"sz,attr,omitempty"`
	Space         *int             `xml:"spc,attr,omitempty"`
	Strike        string           `xml:"strike,attr,omitempty"`
	SolidFill     *DecodeSolidFill `xml:"solidFill,omitempty"`
	Latin         *Font            `xml:"latin,omitempty"`
	EastAsian     *Font            `xml:"ea,omitempty"`
	ComplexScript *Font            `xml:"cs,omitempty"`
}

type DecodeSolidFill struct {
	SolidRGBColor *SolidRGBColor `xml:"srgbClr"`
}

type SolidRGBColor struct {
	Val string `xml:"val,attr"`
}

// Font defines a font typeface element (a:latin, a:ea or a:cs) of the run
// properties.
type Font struct {
	Typeface string `xml:"typeface,attr"`
}
