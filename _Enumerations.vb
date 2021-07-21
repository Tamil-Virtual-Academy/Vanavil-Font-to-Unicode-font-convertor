Imports System.ComponentModel

Namespace Novacode

  Public Enum ListItemType
	Bulleted
	Numbered
  End Enum

  Public Enum SectionBreakType
	defaultNextPage
	evenPage
	oddPage
	continuous
  End Enum


  Public Enum ContainerType
	None
	TOC
	Section
	Cell
	Table
	Header
	Footer
	Paragraph
	Body
  End Enum

  Public Enum PageNumberFormat
	normal
	roman
  End Enum

  Public Enum BorderSize
	one
	two
	three
	four
	five
	six
	seven
	eight
	nine
  End Enum

  Public Enum EditRestrictions
	none
	[readOnly]
	forms
	comments
	trackedChanges
  End Enum

  ''' <summary>
  ''' Table Cell Border styles
  ''' Added by lckuiper @ 20101117
  ''' source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
  ''' </summary>
  Public Enum BorderStyle
	Tcbs_none = 0
	Tcbs_single
	Tcbs_thick
	Tcbs_double
	Tcbs_dotted
	Tcbs_dashed
	Tcbs_dotDash
	Tcbs_dotDotDash
	Tcbs_triple
	Tcbs_thinThickSmallGap
	Tcbs_thickThinSmallGap
	Tcbs_thinThickThinSmallGap
	Tcbs_thinThickMediumGap
	Tcbs_thickThinMediumGap
	Tcbs_thinThickThinMediumGap
	Tcbs_thinThickLargeGap
	Tcbs_thickThinLargeGap
	Tcbs_thinThickThinLargeGap
	Tcbs_wave
	Tcbs_doubleWave
	Tcbs_dashSmallGap
	Tcbs_dashDotStroked
	Tcbs_threeDEmboss
	Tcbs_threeDEngrave
	Tcbs_outset
	Tcbs_inset
	Tcbs_nil
  End Enum

  ''' <summary>
  ''' Table Cell Border Types
  ''' Added by lckuiper @ 20101117
  ''' source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
  ''' </summary>
  Public Enum TableCellBorderType
	Top
	Bottom
	Left
	Right
	InsideH
	InsideV
	TopLeftToBottomRight
	TopRightToBottomLeft
  End Enum

  ''' <summary>
  ''' Table Border Types
  ''' Added by lckuiper @ 20101117
  ''' source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tableborders.aspx
  ''' </summary>
  Public Enum TableBorderType
	Top
	Bottom
	Left
	Right
	InsideH
	InsideV
  End Enum

  ' Patch 7398 added by lckuiper on Nov 16th 2010 @ 2:23 PM
  Public Enum VerticalAlignment
	Top
	Center
	Bottom
  End Enum

  Public Enum Orientation
	Portrait
	Landscape
  End Enum

  Public Enum XmlDocument
	Main
	HeaderOdd
	HeaderEven
	HeaderFirst
	FooterOdd
	FooterEven
	FooterFirst
  End Enum

  Public Enum MatchFormattingOptions
	ExactMatch
	SubsetMatch
  End Enum

  Public Enum Script
	superscript
	subscript
	none
  End Enum

  Public Enum Highlight
	yellow
	green
	cyan
	magenta
	blue
	red
	darkBlue
	darkCyan
	darkGreen
	darkMagenta
	darkRed
	darkYellow
	darkGray
	lightGray
	black
	none
  End Enum

  Public Enum UnderlineStyle
	  none = 0
	  singleLine = 1
	  words = 2
	  doubleLine = 3
	  dotted = 4
	  thick = 6
	  dash = 7
	  dotDash = 9
	  dotDotDash = 10
	  wave = 11
	  dottedHeavy = 20
	  dashedHeavy = 23
	  dashDotHeavy = 25
	  dashDotDotHeavy = 26
	  dashLongHeavy = 27
	  dashLong = 39
	  wavyDouble = 43
	  wavyHeavy = 55
  End Enum

  Public Enum StrikeThrough
	none
	strike
	doubleStrike
  End Enum

  Public Enum Misc
	none
	shadow
	outline
	outlineShadow
	emboss
	engrave
  End Enum

  ''' <summary>
  ''' Change the caps style of text, for use with Append and AppendLine.
  ''' </summary>
  Public Enum CapsStyle
	''' <summary>
	''' No caps, make all characters are lowercase.
	''' </summary>
	none

	''' <summary>
	''' All caps, make every character uppercase.
	''' </summary>
	caps

	''' <summary>
	''' Small caps, make all characters capital but with a small font size.
	''' </summary>
	smallCaps
  End Enum

  ''' <summary>
  ''' Designs\Styles that can be applied to a table.
  ''' </summary>
  Public Enum TableDesign
	Custom
	TableNormal
	TableGrid
	LightShading
	LightShadingAccent1
	LightShadingAccent2
	LightShadingAccent3
	LightShadingAccent4
	LightShadingAccent5
	LightShadingAccent6
	LightList
	LightListAccent1
	LightListAccent2
	LightListAccent3
	LightListAccent4
	LightListAccent5
	LightListAccent6
	LightGrid
	LightGridAccent1
	LightGridAccent2
	LightGridAccent3
	LightGridAccent4
	LightGridAccent5
	LightGridAccent6
	MediumShading1
	MediumShading1Accent1
	MediumShading1Accent2
	MediumShading1Accent3
	MediumShading1Accent4
	MediumShading1Accent5
	MediumShading1Accent6
	MediumShading2
	MediumShading2Accent1
	MediumShading2Accent2
	MediumShading2Accent3
	MediumShading2Accent4
	MediumShading2Accent5
	MediumShading2Accent6
	MediumList1
	MediumList1Accent1
	MediumList1Accent2
	MediumList1Accent3
	MediumList1Accent4
	MediumList1Accent5
	MediumList1Accent6
	MediumList2
	MediumList2Accent1
	MediumList2Accent2
	MediumList2Accent3
	MediumList2Accent4
	MediumList2Accent5
	MediumList2Accent6
	MediumGrid1
	MediumGrid1Accent1
	MediumGrid1Accent2
	MediumGrid1Accent3
	MediumGrid1Accent4
	MediumGrid1Accent5
	MediumGrid1Accent6
	MediumGrid2
	MediumGrid2Accent1
	MediumGrid2Accent2
	MediumGrid2Accent3
	MediumGrid2Accent4
	MediumGrid2Accent5
	MediumGrid2Accent6
	MediumGrid3
	MediumGrid3Accent1
	MediumGrid3Accent2
	MediumGrid3Accent3
	MediumGrid3Accent4
	MediumGrid3Accent5
	MediumGrid3Accent6
	DarkList
	DarkListAccent1
	DarkListAccent2
	DarkListAccent3
	DarkListAccent4
	DarkListAccent5
	DarkListAccent6
	ColorfulShading
	ColorfulShadingAccent1
	ColorfulShadingAccent2
	ColorfulShadingAccent3
	ColorfulShadingAccent4
	ColorfulShadingAccent5
	ColorfulShadingAccent6
	ColorfulList
	ColorfulListAccent1
	ColorfulListAccent2
	ColorfulListAccent3
	ColorfulListAccent4
	ColorfulListAccent5
	ColorfulListAccent6
	ColorfulGrid
	ColorfulGridAccent1
	ColorfulGridAccent2
	ColorfulGridAccent3
	ColorfulGridAccent4
	ColorfulGridAccent5
	ColorfulGridAccent6
	None
  End Enum

  ''' <summary>
  ''' How a Table should auto resize.
  ''' </summary>
  Public Enum AutoFit
	''' <summary>
	''' Autofit to Table contents.
	''' </summary>
	Contents

	''' <summary>
	''' Autofit to Window.
	''' </summary>
	Window

	''' <summary>
	''' Autofit to Column width.
	''' </summary>
	ColumnWidth
	''' <summary>
	'''  Autofit to Fixed column width
	''' </summary>
	Fixed
  End Enum

  Public Enum RectangleShapes
	rect
	roundRect
	snip1Rect
	snip2SameRect
	snip2DiagRect
	snipRoundRect
	round1Rect
	round2SameRect
	round2DiagRect
  End Enum

  Public Enum BasicShapes
	ellipse
	triangle
	rtTriangle
	parallelogram
	trapezoid
	diamond
	pentagon
	hexagon
	heptagon
	octagon
	decagon
	dodecagon
	pie
	chord
	teardrop
	frame
	halfFrame
	corner
	diagStripe
	plus
	plaque
	can
	cube
	bevel
	donut
	noSmoking
	blockArc
	foldedCorner
	smileyFace
	heart
	lightningBolt
	sun
	moon
	cloud
	arc
	backetPair
	bracePair
	leftBracket
	rightBracket
	leftBrace
	rightBrace
  End Enum

  Public Enum BlockArrowShapes
	rightArrow
	leftArrow
	upArrow
	downArrow
	leftRightArrow
	upDownArrow
	quadArrow
	leftRightUpArrow
	bentArrow
	uturnArrow
	leftUpArrow
	bentUpArrow
	curvedRightArrow
	curvedLeftArrow
	curvedUpArrow
	curvedDownArrow
	stripedRightArrow
	notchedRightArrow
	homePlate
	chevron
	rightArrowCallout
	downArrowCallout
	leftArrowCallout
	upArrowCallout
	leftRightArrowCallout
	quadArrowCallout
	circularArrow
  End Enum

  Public Enum EquationShapes
	mathPlus
	mathMinus
	mathMultiply
	mathDivide
	mathEqual
	mathNotEqual
  End Enum

  Public Enum FlowchartShapes
	flowChartProcess
	flowChartAlternateProcess
	flowChartDecision
	flowChartInputOutput
	flowChartPredefinedProcess
	flowChartInternalStorage
	flowChartDocument
	flowChartMultidocument
	flowChartTerminator
	flowChartPreparation
	flowChartManualInput
	flowChartManualOperation
	flowChartConnector
	flowChartOffpageConnector
	flowChartPunchedCard
	flowChartPunchedTape
	flowChartSummingJunction
	flowChartOr
	flowChartCollate
	flowChartSort
	flowChartExtract
	flowChartMerge
	flowChartOnlineStorage
	flowChartDelay
	flowChartMagneticTape
	flowChartMagneticDisk
	flowChartMagneticDrum
	flowChartDisplay
  End Enum

  Public Enum StarAndBannerShapes
	irregularSeal1
	irregularSeal2
	star4
	star5
	star6
	star7
	star8
	star10
	star12
	star16
	star24
	star32
	ribbon
	ribbon2
	ellipseRibbon
	ellipseRibbon2
	verticalScroll
	horizontalScroll
	wave
	doubleWave
  End Enum

  Public Enum CalloutShapes
	wedgeRectCallout
	wedgeRoundRectCallout
	wedgeEllipseCallout
	cloudCallout
	borderCallout1
	borderCallout2
	borderCallout3
	accentCallout1
	accentCallout2
	accentCallout3
	callout1
	callout2
	callout3
	accentBorderCallout1
	accentBorderCallout2
	accentBorderCallout3
  End Enum

  ''' <summary>
  ''' Text alignment of a Paragraph.
  ''' </summary>
  Public Enum Alignment
	''' <summary>
	''' Align Paragraph to the left.
	''' </summary>
	left

	''' <summary>
	''' Align Paragraph as centered.
	''' </summary>
	center

	''' <summary>
	''' Align Paragraph to the right.
	''' </summary>
	right

	''' <summary>
	''' (Justified) Align Paragraph to both the left and right margins, adding extra space between content as necessary.
	''' </summary>
	both
  End Enum

  Public Enum Direction
	LeftToRight
	RightToLeft
  End Enum

  ''' <summary>
  ''' Paragraph edit types
  ''' </summary>
  Friend Enum EditType
	''' <summary>
	''' A ins is a tracked insertion
	''' </summary>
	ins

	''' <summary>
	''' A del is  tracked deletion
	''' </summary>
	del
  End Enum

  ''' <summary>
  ''' Custom property types.
  ''' </summary>
  Friend Enum CustomPropertyType
	''' <summary>
	''' System.String
	''' </summary>
	Text

	''' <summary>
	''' System.DateTime
	''' </summary>
	[Date]

	''' <summary>
	''' System.Int32
	''' </summary>
	NumberInteger

	''' <summary>
	''' System.Double
	''' </summary>
	NumberDecimal

	''' <summary>
	''' System.Boolean
	''' </summary>
	YesOrNo
  End Enum

  ''' <summary>
  ''' Text types in a Run
  ''' </summary>
  Public Enum RunTextType
	''' <summary>
	''' System.String
	''' </summary>
	Text

	''' <summary>
	''' System.String
	''' </summary>
	DelText
  End Enum
  Public Enum LineSpacingType
	  Line
	  Before
	  After
  End Enum

  Public Enum LineSpacingTypeAuto
	  AutoBefore
	  AutoAfter
	  [Auto]
	  None
  End Enum

  ''' <summary>
  ''' Cell margin for all sides of the table cell.
  ''' </summary>
  Public Enum TableCellMarginType
	  ''' <summary>
	  ''' The left cell margin.
	  ''' </summary>
	  left
	  ''' <summary>
	  ''' The right cell margin.
	  ''' </summary>
	  right
	  ''' <summary>
	  ''' The bottom cell margin.
	  ''' </summary>
	  bottom
	  ''' <summary>
	  ''' The top cell margin.
	  ''' </summary>
	  top
  End Enum

  Public Enum HeadingType
	  <Description("Heading1")>
	  Heading1

	  <Description("Heading2")>
	  Heading2

	  <Description("Heading3")>
	  Heading3

	  <Description("Heading4")>
	  Heading4

	  <Description("Heading5")>
	  Heading5

	  <Description("Heading6")>
	  Heading6

	  <Description("Heading7")>
	  Heading7

	  <Description("Heading8")>
	  Heading8

	  <Description("Heading9")>
	  Heading9


'      		 
'       * The Character Based Headings below do not work in the same way as the headings 1-9 above, but appear on the same list in word. 
'       * I have kept them here for reference in case somebody else things its just a matter of adding them in to gain extra headings
'       
	  #Region "Other character (NOT paragraph) based Headings"
	  '[Description("NoSpacing")]
	  'NoSpacing,

	  '[Description("Title")]
	  'Title,

	  '[Description("Subtitle")]
	  'Subtitle,

	  '[Description("Quote")]
	  'Quote,

	  '[Description("IntenseQuote")]
	  'IntenseQuote,

	  '[Description("Emphasis")]
	  'Emphasis,

	  '[Description("IntenseEmphasis")]
	  'IntenseEmphasis,

	  '[Description("Strong")]
	  'Strong,

	  '[Description("ListParagraph")]
	  'ListParagraph,

	  '[Description("SubtleReference")]
	  'SubtleReference,

	  '[Description("IntenseReference")]
	  'IntenseReference,

	  '[Description("BookTitle")]
	  'BookTitle, 
	  #End Region


  End Enum
  Public Enum TextDirection
	  btLr
	  right
  End Enum

	''' <summary>
	''' Represents the switches set on a TOC.
	''' </summary>
	<Flags>
	Public Enum TableOfContentsSwitches
		None = 0 << 0
		<Description("\a")>
		A = 1 << 0
		<Description("\b")>
		B = 1 << 1
		<Description("\c")>
		C = 1 << 2
		<Description("\d")>
		D = 1 << 3
		<Description("\f")>
		F = 1 << 4
		<Description("\h")>
		H = 1 << 5
		<Description("\l")>
		L = 1 << 6
		<Description("\n")>
		N = 1 << 7
		<Description("\o")>
		O = 1 << 8
		<Description("\p")>
		P = 1 << 9
		<Description("\s")>
		S = 1 << 10
		<Description("\t")>
		T = 1 << 11
		<Description("\u")>
		U = 1 << 12
		<Description("\w")>
		W = 1 << 13
		<Description("\x")>
		X = 1 << 14
		<Description("\z")>
		Z = 1 << 15
	End Enum

End Namespace