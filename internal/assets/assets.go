package assets

import _ "embed"

//go:embed slide.xml.tpl
var MSSlideTpl string

//go:embed ninelock.png
var MSMarkImage []byte

//go:embed drawing.xml
var MSDrawing []byte

//go:embed drawing.xml.tpl
var MSDrawingTpl string
